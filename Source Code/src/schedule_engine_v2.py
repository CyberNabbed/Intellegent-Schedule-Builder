from ortools.sat.python import cp_model
import calendar
from datetime import date
import math

# Shift Definitions (Indices)
# Phone Shifts
P1 = 0 # 7:30 - 10:00 (Early Start Only)
P2 = 1 # 10:00 - 12:00
P3 = 2 # 12:00 - 2:30
P4 = 3 # 2:30 - 5:00 (Late Start Only)

# Front Desk Shifts
FD1 = 4 # 8:30 - 12:00
FD2 = 5 # 12:00 - 4:00

ALL_SHIFTS = [P1, P2, P3, P4, FD1, FD2]
PHONE_SHIFTS = [P1, P2, P3, P4]
FD_SHIFTS = [FD1, FD2]

SHIFT_NAMES = {
    P1: "Phone (7:30-10)",
    P2: "Phone (10-12)",
    P3: "Phone (12-2:30)",
    P4: "Phone (2:30-5)",
    FD1: "Front Desk (8:30-12)",
    FD2: "Front Desk (12-4)"
}

# Shift start/end times as (start_hour, start_min, end_hour, end_min)
# Used by ICS export to create calendar events with correct timestamps
SHIFT_TIMES = {
    P1:  (7,  30, 10,  0),   # P1:  7:30 AM – 10:00 AM
    P2:  (10,  0, 12,  0),   # P2: 10:00 AM – 12:00 PM
    P3:  (12,  0, 14, 30),   # P3: 12:00 PM –  2:30 PM
    P4:  (14, 30, 17,  0),   # P4:  2:30 PM –  5:00 PM
    FD1: (8,  30, 12,  0),   # FD1: 8:30 AM – 12:00 PM
    FD2: (12,  0, 16,  0),   # FD2: 12:00 PM – 4:00 PM
}

class ScheduleGeneratorV2:
    def __init__(self, employees, year, month, skip_weekends=True, holidays=None,
                 start_from_day=None, generate_full_weeks=False):
        """
        employees: List of dicts [{"name": "Name", "type": "early" | "late"}]
        holidays: List of integers (day of month) to skip coverage.
        start_from_day: Start generating from this day of the month (optional)
        generate_full_weeks: If True, extend schedule to complete calendar weeks
        """
        self.employees = employees
        self.year = year
        self.month = month
        self.skip_weekends = skip_weekends
        self.holidays = holidays if holidays else []
        self.start_from_day = start_from_day
        self.generate_full_weeks = generate_full_weeks
        self.dates = self._generate_dates()
        self.num_days = len(self.dates)

        # Validate that there are working days in the selected month
        if self.num_days == 0:
            month_name = calendar.month_name[month]
            raise ValueError(f"No working days found in {month_name} {year}. "
                           "All days are either weekends or holidays.")

        self.model = cp_model.CpModel()
        self.shifts = {}  # (employee_name, day_index, shift_id) -> BoolVar
        self.solver = cp_model.CpSolver()
        self.status = None

    def is_holiday(self, day_index):
        """Check if a given day index is a holiday in the target month."""
        if day_index < 0 or day_index >= self.num_days:
            return False
        dt = self.dates[day_index]
        return dt.day in self.holidays and dt.month == self.month

    def _generate_dates(self):
        """
        Generates a list of date objects for scheduling, with support for:
        - Starting from a specific day of the month (start_from_day)
        - Extending to complete full calendar weeks (generate_full_weeks)
        - Skipping weekends
        """
        from datetime import timedelta

        num_days_in_month = calendar.monthrange(self.year, self.month)[1]
        valid_dates = []

        # Determine the start day
        start_day = self.start_from_day if self.start_from_day else 1

        # Generate dates from start_day through end of month
        for day in range(start_day, num_days_in_month + 1):
            dt = date(self.year, self.month, day)
            # weekday(): 0=Mon, 6=Sun. Skip if >= 5 (Sat/Sun)
            if self.skip_weekends and dt.weekday() >= 5:
                continue
            valid_dates.append(dt)

        # If generate_full_weeks is enabled, extend to complete the last week
        if self.generate_full_weeks and valid_dates:
            last_date = valid_dates[-1]
            last_weekday = last_date.weekday()  # 0=Mon, 4=Fri

            # If the last date is not a Friday, extend to Friday
            if last_weekday < 4:  # 0-3 = Mon-Thu
                days_to_add = 4 - last_weekday
                for i in range(1, days_to_add + 1):
                    next_date = last_date + timedelta(days=i)
                    if self.skip_weekends and next_date.weekday() >= 5:
                        continue
                    valid_dates.append(next_date)

        return valid_dates

    def diagnose(self, time_off_requests):
        """
        Analyzes constraints to find impossible days.
        time_off_requests: dict {(name, day_index): True}
        """
        issues = []
        staff_names = [e["name"] for e in self.employees]
        group_a = [e["name"] for e in self.employees if e["type"] == "early"]
        group_b = [e["name"] for e in self.employees if e["type"] == "late"]

        for d in range(self.num_days):
            # Only skip holidays in the target month (not spillover into next month)
            if self.dates[d].day in self.holidays and self.dates[d].month == self.month:
                continue

            date_str = self.dates[d].strftime("%b %d")
            
            # Who is available today?
            available = []
            available_a = []
            available_b = []
            
            for emp in staff_names:
                if (emp, d) not in time_off_requests:
                    available.append(emp)
                    if emp in group_a: available_a.append(emp)
                    if emp in group_b: available_b.append(emp)
            
            # Check 1: Absolute Minimum for Phones (Need 4 unique people)
            if len(available) < 4:
                issues.append(f"{date_str}: Only {len(available)} staff available. Need 4 minimum for phones.")
                continue

            # Check 2: Morning Coverage (Need at least 1 Group A)
            if len(available_a) == 0:
                issues.append(f"{date_str}: No 'Early' staff available. Cannot cover 7:30 AM.")
            
            # Check 3: Afternoon Coverage (Need at least 1 Group B)
            if len(available_b) == 0:
                issues.append(f"{date_str}: No 'Late' staff available. Cannot cover 2:30 PM.")
                
        return issues

    def build_model(self, time_off_requests=None, soft_coverage=False):
        """
        Build the constraint programming model.

        time_off_requests: dict {(name, day_index): True} - days when employees are unavailable
        soft_coverage: if True, allows partial coverage when constraints are too tight
        """
        if time_off_requests is None:
            time_off_requests = {}

        staff_names = [e["name"] for e in self.employees]

        # Calculate available days per person (for proportional fairness)
        available_days = {}
        for emp in staff_names:
            count = 0
            for d in range(self.num_days):
                # Only count holidays in the target month
                is_holiday = self.dates[d].day in self.holidays and self.dates[d].month == self.month
                if not is_holiday and (emp, d) not in time_off_requests:
                    count += 1
            available_days[emp] = count

        # 1. Create Variables
        for emp in staff_names:
            for d in range(self.num_days):
                for s in ALL_SHIFTS:
                    self.shifts[(emp, d, s)] = self.model.NewBoolVar(f"shift_{emp}_d{d}_s{s}")

        # 2. Daily Coverage
        total_assigned_shifts = []

        for d in range(self.num_days):
            # Only treat as holiday if in the target month
            is_holiday = self.dates[d].day in self.holidays and self.dates[d].month == self.month

            if is_holiday:
                # Force everyone OFF on holidays
                for emp in staff_names:
                    for s in ALL_SHIFTS:
                        self.model.Add(self.shifts[(emp, d, s)] == 0)
                continue # Skip coverage constraints

            for s in PHONE_SHIFTS:
                coverage = sum(self.shifts[(emp, d, s)] for emp in staff_names)
                if soft_coverage:
                    # Allow 0 or 1, prefer 1
                    self.model.Add(coverage <= 1)
                    total_assigned_shifts.append(coverage)
                else:
                    # Strict: Must be exactly 1
                    self.model.Add(coverage == 1)
            
            for s in FD_SHIFTS:
                coverage = sum(self.shifts[(emp, d, s)] for emp in staff_names)
                if soft_coverage:
                    self.model.Add(coverage <= 1)
                    total_assigned_shifts.append(coverage)
                else:
                    self.model.Add(coverage == 1)
        
        if soft_coverage:
            # Maximize the number of filled slots
            self.model.Maximize(sum(total_assigned_shifts))

        # 3. Time Overlaps (Physical Impossibility)
        # P1 (7:30-10) overlaps FD1 (8:30-12)
        # P2 (10-12) overlaps FD1 (8:30-12)
        # P3 (12-2:30) overlaps FD2 (12-4)
        # P4 (2:30-5) overlaps FD2 (12-4)
        for emp in staff_names:
            for d in range(self.num_days):
                self.model.Add(self.shifts[(emp, d, P1)] + self.shifts[(emp, d, FD1)] <= 1)
                self.model.Add(self.shifts[(emp, d, P2)] + self.shifts[(emp, d, FD1)] <= 1)
                self.model.Add(self.shifts[(emp, d, P3)] + self.shifts[(emp, d, FD2)] <= 1)
                self.model.Add(self.shifts[(emp, d, P4)] + self.shifts[(emp, d, FD2)] <= 1)

        # 4. Max 1 Phone Shift per Person per Day (Hard Constraint)
        for emp in staff_names:
            for d in range(self.num_days):
                self.model.Add(sum(self.shifts[(emp, d, s)] for s in PHONE_SHIFTS) <= 1)
                # NEW: Max 1 Front Desk Shift per Person per Day (Hard Constraint)
                self.model.Add(sum(self.shifts[(emp, d, s)] for s in FD_SHIFTS) <= 1)

        # 5. Type-Based Restrictions (Hard Constraint)
        for emp_data in self.employees:
            name = emp_data["name"]
            shift_type = emp_data["type"]
            
            if shift_type == "late":
                # Late starters (8:30) cannot work P1 (7:30 start)
                for d in range(self.num_days):
                    self.model.Add(self.shifts[(name, d, P1)] == 0)
            
            elif shift_type == "early":
                # Early starters (leave 4:00) cannot work P4 (ends 5:00)
                for d in range(self.num_days):
                    self.model.Add(self.shifts[(name, d, P4)] == 0)

        # 6. Equity Rule (Updated: Proportional to Availability)
        # NEW: Fairness based on available days, not total team size
        # People who take time off work proportionally fewer shifts

        # Calculate working days (excluding holidays in target month only)
        num_working_days = sum(
            1 for d in range(self.num_days)
            if not (self.dates[d].day in self.holidays and self.dates[d].month == self.month)
        )

        total_phone_slots = num_working_days * 4
        total_fd_slots = num_working_days * 2

        # Calculate total person-days available across the entire team
        total_person_days = sum(available_days[emp] for emp in staff_names)

        if total_person_days == 0:
            raise ValueError("No staff available for scheduling (all days are time-off or holidays)")

        # Calculate fair rate per person-day
        # Example: 80 phone slots / 110 total person-days = 0.727 phone shifts per person-day
        phone_rate = total_phone_slots / total_person_days
        fd_rate = total_fd_slots / total_person_days

        for emp in staff_names:
            days_avail = available_days[emp]

            # Calculate this person's fair share based on their availability
            # Example: Person available 15 days → 15 * 0.727 = 10.9 phone shifts
            target_phone = phone_rate * days_avail
            target_fd = fd_rate * days_avail

            # Allow +/- 1 flexibility (floor to ceil)
            min_p = math.floor(target_phone)
            max_p = math.ceil(target_phone)

            min_fd = math.floor(target_fd)
            max_fd = math.ceil(target_fd)

            # Count total phone shifts for this person
            p_count = sum(self.shifts[(emp, d, s)] for d in range(self.num_days) for s in PHONE_SHIFTS)
            # Proportional bounds based on their availability
            self.model.Add(p_count >= int(min_p))
            self.model.Add(p_count <= int(max_p))

            # Count total FD shifts for this person
            fd_count = sum(self.shifts[(emp, d, s)] for d in range(self.num_days) for s in FD_SHIFTS)
            self.model.Add(fd_count >= int(min_fd))
            self.model.Add(fd_count <= int(max_fd))

        # 7. Soft Constraints (Preferences)
        penalties = []

        # Goal: Variety (Maximize changes in daily schedule)
        for emp in staff_names:
            for d in range(self.num_days - 1):
                for s in ALL_SHIFTS:
                    # If worked shift S on day D and D+1, penalty
                    both_days = self.model.NewBoolVar(f"repeat_{emp}_{d}_{s}")
                    self.model.AddBoolAnd([self.shifts[(emp, d, s)], self.shifts[(emp, d+1, s)]]).OnlyEnforceIf(both_days)
                    self.model.AddBoolOr([self.shifts[(emp, d, s)].Not(), self.shifts[(emp, d+1, s)].Not()]).OnlyEnforceIf(both_days.Not())
                    penalties.append(both_days)

        # NEW: Goal: Spread Front Desk shifts (Avoid Consecutive FD Days)
        for emp in staff_names:
            for d in range(self.num_days - 1):
                # Is working FD today?
                fd_today = self.model.NewBoolVar(f"fd_today_{emp}_{d}")
                self.model.Add(sum(self.shifts[(emp, d, s)] for s in FD_SHIFTS) >= 1).OnlyEnforceIf(fd_today)
                self.model.Add(sum(self.shifts[(emp, d, s)] for s in FD_SHIFTS) == 0).OnlyEnforceIf(fd_today.Not())

                # Is working FD tomorrow?
                fd_tomorrow = self.model.NewBoolVar(f"fd_tomorrow_{emp}_{d}")
                self.model.Add(sum(self.shifts[(emp, d+1, s)] for s in FD_SHIFTS) >= 1).OnlyEnforceIf(fd_tomorrow)
                self.model.Add(sum(self.shifts[(emp, d+1, s)] for s in FD_SHIFTS) == 0).OnlyEnforceIf(fd_tomorrow.Not())

                # Penalty if BOTH are true
                consecutive_fd = self.model.NewBoolVar(f"consecutive_fd_{emp}_{d}")
                self.model.AddBoolAnd([fd_today, fd_tomorrow]).OnlyEnforceIf(consecutive_fd)
                self.model.AddBoolOr([fd_today.Not(), fd_tomorrow.Not()]).OnlyEnforceIf(consecutive_fd.Not())

                # Weight this decently high (3)
                penalties.append(consecutive_fd)
                penalties.append(consecutive_fd)
                penalties.append(consecutive_fd)

        # Goal: Minimize "Double Duty" (Phone + FD in same day)
        # AND distribute it fairly across team members
        double_duty_vars = {}  # {emp: [is_double_day0, is_double_day1, ...]}

        for emp in staff_names:
            double_duty_vars[emp] = []
            for d in range(self.num_days):
                is_phone = self.model.NewBoolVar(f"is_phone_{emp}_{d}")
                is_fd = self.model.NewBoolVar(f"is_fd_{emp}_{d}")

                self.model.Add(sum(self.shifts[(emp, d, s)] for s in PHONE_SHIFTS) >= 1).OnlyEnforceIf(is_phone)
                self.model.Add(sum(self.shifts[(emp, d, s)] for s in PHONE_SHIFTS) == 0).OnlyEnforceIf(is_phone.Not())

                self.model.Add(sum(self.shifts[(emp, d, s)] for s in FD_SHIFTS) >= 1).OnlyEnforceIf(is_fd)
                self.model.Add(sum(self.shifts[(emp, d, s)] for s in FD_SHIFTS) == 0).OnlyEnforceIf(is_fd.Not())

                # If both are true, penalty
                is_double = self.model.NewBoolVar(f"double_{emp}_{d}")
                self.model.AddBoolAnd([is_phone, is_fd]).OnlyEnforceIf(is_double)
                self.model.AddBoolOr([is_phone.Not(), is_fd.Not()]).OnlyEnforceIf(is_double.Not())

                double_duty_vars[emp].append(is_double)

                # Weight double duty occurrence (5x penalty)
                penalties.append(is_double)
                penalties.append(is_double)
                penalties.append(is_double)
                penalties.append(is_double)
                penalties.append(is_double)

        # NEW: Goal: Minimize Back-to-Back Shifts (e.g., P2 -> FD2 or FD1 -> P3)
        for emp in staff_names:
            for d in range(self.num_days):
                # P2 (10-12) and FD2 (12-4) touch at 12:00
                b2b_1 = self.model.NewBoolVar(f"b2b_1_{emp}_{d}")
                self.model.AddBoolAnd([self.shifts[(emp, d, P2)], self.shifts[(emp, d, FD2)]]).OnlyEnforceIf(b2b_1)
                self.model.AddBoolOr([self.shifts[(emp, d, P2)].Not(), self.shifts[(emp, d, FD2)].Not()]).OnlyEnforceIf(b2b_1.Not())

                # FD1 (8:30-12) and P3 (12-2:30) touch at 12:00
                b2b_2 = self.model.NewBoolVar(f"b2b_2_{emp}_{d}")
                self.model.AddBoolAnd([self.shifts[(emp, d, FD1)], self.shifts[(emp, d, P3)]]).OnlyEnforceIf(b2b_2)
                self.model.AddBoolOr([self.shifts[(emp, d, FD1)].Not(), self.shifts[(emp, d, P3)].Not()]).OnlyEnforceIf(b2b_2.Not())

                # Add extra heavy penalties for these (weight: 4x)
                # Since these are also "double duty", total penalty for back-to-back = 9x
                for _ in range(4):
                    penalties.append(b2b_1)
                    penalties.append(b2b_2)

        # NEW: Fair Distribution of Double Duty
        # Ensure no one gets stuck with all the double duty - distribute fairly
        # Count total double duty days per person
        double_duty_counts = {}
        for emp in staff_names:
            double_duty_counts[emp] = sum(double_duty_vars[emp])

        # Penalize imbalance: for each pair of employees, penalize if their
        # double duty counts differ by more than 1
        # This encourages round-robin distribution
        for i, emp1 in enumerate(staff_names):
            for emp2 in staff_names[i+1:]:
                # Create variable for difference > 1
                diff_var = self.model.NewIntVar(0, self.num_days, f"dd_diff_{emp1}_{emp2}")
                self.model.Add(diff_var >= double_duty_counts[emp1] - double_duty_counts[emp2] - 1)
                self.model.Add(diff_var >= double_duty_counts[emp2] - double_duty_counts[emp1] - 1)

                # Add penalty for each unit of imbalance (weight: 3x)
                # Less than the 5x for having double duty, but significant
                penalties.append(diff_var)
                penalties.append(diff_var)
                penalties.append(diff_var)
        
        if soft_coverage:
            # Objective: Maximize coverage (primary), Minimize penalties (secondary)
            # Coverage weight: 1000, Penalty weight: 1
            # Since penalties are minimized, we subtract them from the maximization objective.
            assigned_sum = sum(total_assigned_shifts) if total_assigned_shifts else 0
            penalty_sum = sum(penalties) if penalties else 0
            self.model.Maximize((assigned_sum * 1000) - penalty_sum)
        else:
            self.model.Minimize(sum(penalties))

    def solve(self):
        self.status = self.solver.Solve(self.model)
        if self.status == cp_model.OPTIMAL or self.status == cp_model.FEASIBLE:
            return True
        return False

    def export_text(self):
        if self.status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            return "No solution found."

        staff_names = [e["name"] for e in self.employees]
        output = []
        output.append(f"Schedule for {calendar.month_name[self.month]} {self.year}")
        output.append(f"Total Working Days: {self.num_days}")
        output.append("-" * 40)
        
        for d in range(self.num_days):
            date_str = self.dates[d].strftime("%a %b %d")
            output.append(f"\n{date_str}:")
            # Group by Shift Type
            phone_workers = []
            for s in PHONE_SHIFTS:
                for emp in staff_names:
                    if self.solver.Value(self.shifts[(emp, d, s)]):
                        phone_workers.append(f"{SHIFT_NAMES[s]}: {emp}")
            
            fd_workers = []
            for s in FD_SHIFTS:
                for emp in staff_names:
                    if self.solver.Value(self.shifts[(emp, d, s)]):
                        fd_workers.append(f"{SHIFT_NAMES[s]}: {emp}")
            
            for line in phone_workers: output.append(f"  {line}")
            for line in fd_workers: output.append(f"  {line}")

        # Stats
        output.append("\n--- Stats (Equity Check) ---")
        output.append(f"{'Name':<10} | {'Phone':<5} | {'FD':<5} | {'Total':<5}")
        output.append("-" * 30)
        
        for emp in staff_names:
            p_count = sum(self.solver.Value(self.shifts[(emp, d, s)]) for d in range(self.num_days) for s in PHONE_SHIFTS)
            fd_count = sum(self.solver.Value(self.shifts[(emp, d, s)]) for d in range(self.num_days) for s in FD_SHIFTS)
            output.append(f"{emp:<10} | {p_count:<5} | {fd_count:<5} | {p_count + fd_count:<5}")
            
        return "\n".join(output)

    def export_as_table(self):
        """Generates an ASCII-formatted table similar to the Excel export."""
        if self.status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            return "No solution found."

        weeks = self.get_weekly_matrix()
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        
        output = []
        output.append(f"SCHEDULE FOR {calendar.month_name[self.month].upper()} {self.year}\n")

        for week_idx, week_data in enumerate(weeks):
            output.append(f"--- WEEK {week_idx + 1} ---")
            
            # Header Row 1: Day Names
            header1 = f"{'':<20} |"
            for d_name in days:
                header1 += f" {d_name:<15} |"
            output.append(header1)
            
            # Header Row 2: Dates
            header2 = f"{'':<20} |"
            for d_date in week_data["dates"]:
                header2 += f" {d_date:<15} |"
            output.append(header2)
            
            output.append("-" * len(header1))
            
            # Shift Rows
            for s_idx in ALL_SHIFTS:
                shift_name = SHIFT_NAMES[s_idx]
                row = f"{shift_name:<20} |"
                workers = week_data["matrix"][s_idx]
                for worker in workers:
                    row += f" {worker:<15} |"
                output.append(row)
            
            output.append("\n")

        # Fairness Audit
        output.append("--- FAIRNESS AUDIT ---")
        headers = ["Employee", "Phone", "FD", "Total"]
        audit_header = f" {headers[0]:<15} | {headers[1]:<7} | {headers[2]:<7} | {headers[3]:<7}"
        output.append(audit_header)
        output.append("-" * len(audit_header))
        
        stats = self.get_stats()
        for stat in stats:
            row = f" {stat['name']:<15} | {stat['phone']:<7} | {stat['fd']:<7} | {stat['total']:<7}"
            output.append(row)
            
        return "\n".join(output)

    def get_weekly_matrix(self):
        """
        Returns data structured by week for the new Excel layout.
        Structure:
        [
            {
                "dates": ["Oct 1", "Oct 2", ...],  # Labels for columns
                "matrix": {
                    "Phone (7:30-10)": ["Patrick", "Jesus", ...],
                    "Phone (10-12)": [...],
                    ...
                }
            },
            ...
        ]
        """
        weeks = []
        current_week = None
        
        staff_names = [e["name"] for e in self.employees]
        
        for d in range(self.num_days):
            dt = self.dates[d]
            wd = dt.weekday() # 0=Mon, 4=Fri
            
            # Start new week if:
            # 1. No current week
            # 2. It's Monday (wd==0)
            # 3. Current week is "full" (last date was Friday?) - actually just check Mon
            
            if current_week is None or wd == 0 or (d > 0 and dt.weekday() < self.dates[d-1].weekday()):
                current_week = {
                    "dates": [""] * 5,
                    "matrix": {s: [""] * 5 for s in ALL_SHIFTS},
                    "holidays": [False] * 5  # Track which columns are holidays
                }
                weeks.append(current_week)

            # Map valid weekday to 0-4 index (Mon=0..Fri=4)
            # Since we skip weekends, dt.weekday() is 0,1,2,3,4
            col_idx = wd
            if col_idx > 4: continue # Should not happen with skip_weekends=True

            # Set Date Label
            current_week["dates"][col_idx] = dt.strftime("%b %d")

            # Mark if this day is a holiday
            current_week["holidays"][col_idx] = self.is_holiday(d)
            
            # Fill Shifts
            for s in ALL_SHIFTS:
                worker = ""
                for emp in staff_names:
                    if self.solver.Value(self.shifts[(emp, d, s)]):
                        worker = emp
                        break
                current_week["matrix"][s][col_idx] = worker
                
        return weeks

    def get_stats(self):
        """Returns fairness statistics for the schedule."""
        stats = []
        staff_names = [e["name"] for e in self.employees]
        
        for emp in staff_names:
            p_count = sum(self.solver.Value(self.shifts[(emp, d, s)]) for d in range(self.num_days) for s in PHONE_SHIFTS)
            fd_count = sum(self.solver.Value(self.shifts[(emp, d, s)]) for d in range(self.num_days) for s in FD_SHIFTS)
            stats.append({
                "name": emp,
                "phone": p_count,
                "fd": fd_count,
                "total": p_count + fd_count
            })
        return stats