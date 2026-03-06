#     Intellegent Schedule Builder

This is a desktop tool I built to solve a specific variant of the "Nurse Scheduling Problem." It’s designed for helpdesk environments where you have a mix of early/late staff and need to fill specific phone and front desk slots without making the team miserable.

<p align="center">
  <picture>
    <img alt="General UI" src="https://github.com/CyberNabbed/IntellegentTeamScheduler/blob/main/Sample%20Photos/General%20UI.png?raw=true">
  </picture>
</p>

Instead of just randomly picking names, this app uses a constraint optimization engine to find the most "mathematically fair" schedule for everyone involved.

---

### What it actually does

*   **Solve in Milliseconds:** The CP-SAT solver handles a 6-person, 1-month roster almost instantly.
*   **Human-in-the-loop Editing:** You can drag and drop names in the UI to make manual tweaks. The engine runs a background validation to highlight logic errors (like overlapping shifts) in red.
    <p align="center">
      <picture>
        <img alt="Main Pane Settings" src="https://github.com/CyberNabbed/IntellegentTeamScheduler/blob/main/Sample%20Photos/Main%20Pane.png?raw=true">
      </picture>
    </p>
*   **Fairness Audit:** A side panel shows exactly how many shifts each person has. It's an easy way to verify the math is doing its job.
    <p align="center">
      <picture>
        <img alt="Fairness Audit" src="https://github.com/CyberNabbed/IntellegentTeamScheduler/blob/main/Sample%20Photos/Fairness%20Audit.png?raw=true">
      </picture>
    </p>
*   **Calendar & Excel Export:** It spits out a color-coded Excel sheet and individual `.ics` files that team members can double-click to add to their Outlook or Google calendars.
    <p align="center">
      <picture>
        <img alt="Excel Export" src="https://github.com/CyberNabbed/IntellegentTeamScheduler/blob/main/Sample%20Photos/Excel%20Export.png?raw=true">
      </picture>
    </p>
    <p align="center">
      <picture>
        <img alt="Calendar Export Options" src="https://github.com/CyberNabbed/IntellegentTeamScheduler/blob/main/Sample%20Photos/Calendar%20Export%20Options.png?raw=true">
      </picture>
    </p>
*   **Outlook Integration:** On Windows, it can auto-generate a draft email with the schedule attached to save time on distribution.
    <p align="center">
      <picture>
        <img alt="Schedule Emailer" src="https://github.com/CyberNabbed/IntellegentTeamScheduler/blob/main/Sample%20Photos/Schedule%20Emailer.png?raw=true">
      </picture>
    </p>

---

### How it works (The Math)

Most schedulers use simple loops or "greedy" algorithms that fill one slot at a time. The problem with that is you often end up "painted into a corner" where the last person gets stuck with a terrible week.

I used **Google OR-Tools (CP-SAT)** to treat the entire month as a single multi-variable equation. The solver looks at every possible permutation to find the one that satisfies these rules:

*   **Hard Constraints:** These are non-negotiable. No one works two shifts in a day, early-type staff don't work late-type shifts, and every slot is filled exactly as required.
*   **The Fairness Model:** This is the core of the project. The solver calculates the "variance" in shift counts across the team. I added a constraint that the difference in total shifts between any two employees can never be more than 1.
*   **Soft Constraints (Heuristics):** The model assigns "penalty points" for things like "clumping" (putting someone on the Front Desk two days in a row) or "double duty" (phone and desk on the same day). The engine then runs until it finds the schedule with the lowest possible penalty score.

---

### Tech Stack

*   **Python 3.10+**
*   **OR-Tools:** For the constraint programming logic.
*   **CustomTkinter:** For a clean, modern GUI that handles dark/light mode and high-DPI scaling.
*   **OpenPyXL:** For generating the formatted Excel exports.
*   **Pytest:** I've included a suite of tests to make sure the solver logic and session saving don't break during updates.

### Setup & Running

```bash
# Get the dependencies
pip install -r requirements.txt

# Launch the app
python3 src/scheduler.py
```

On the first run, the app will launch a **Setup Wizard** to help you add your team members and define their shift types. You can update this later at any time directly through the UI. 

<p align="center">
  <picture>
    <img alt="Team Management Pane" src="https://github.com/CyberNabbed/IntellegentTeamScheduler/blob/main/Sample%20Photos/Team%20Management%20Pane.png?raw=true">
  </picture>
</p>

---
