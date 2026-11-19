# Classroom Debate Topic Selector & Team Matcher (Google Apps Script)

A Google Sheets Apps Script suite for picking classroom debate topics and assigning students to balanced Pro/Con teams. It runs a two-round workflow inside Google Sheets and Google Forms, with audit trails, tie warnings, and a student satisfaction report.

---

## Quick start (each debate cycle)

> **Forgot how this works?** Open the spreadsheet and click **Debate Tools → ❓ What do I do next?** It shows where you are, links to both forms, response counts, and the next step.

| Step | Menu item | When |
|---|---|---|
| 0 | **Start new debate cycle** | *Before* you share the Round 1 form. Older ballots stop counting. |
| 1 | **Run Borda Count (Round 1)** | After students vote on topics. |
| 2 | **Generate Round 2 Form** | Right after step 1. Share the student link it shows. |
| 3 | **Match Teams (Round 2)** | After students submit Round 2. |

You can run every step from any tab. The script finds the right response tabs on its own.

---

## Features

### Round 1: Borda count consensus voting
- Tallies ranked-choice ballots from a Google Form multiple-choice grid to find the two topics with the broadest support.
- Counts only ballots submitted since the current cycle started, so you can reuse the same form and spreadsheet every time.
- Ignores columns for topics you've removed or renamed on the form. Google Sheets keeps those old columns, but they no longer count.
- Breaks ties alphabetically so results are reproducible, and warns you when a tie decides which topic gets the second winning spot.
- Creates a timestamped **Borda Results** audit tab that highlights the winners.

### Automatic Round 2 form
- Builds a new Google Form with the two winning topics and links its responses to the same spreadsheet.
- Turns on **one response per column** for the topic-ranking grid automatically, so students can't rank both topics "1".
- Asks for side preference separately, on a 5-point scale: Strongly Prefer Pro, Prefer Pro, No Preference, Prefer Con, Strongly Prefer Con.

### Round 2: fair team matching
- Uses a Fisher-Yates shuffle to remove any advantage from submitting early.
- Calculates exact team sizes. An odd extra student goes to the more popular topic or side.
- Uses a preference gradient to balance Pro and Con, so students with no preference or a mild one are moved before anyone with a strong preference.
- If a student submits more than once, only their **latest** submission is used. Names are matched ignoring capitalization and extra spaces.
- Creates a timestamped **Final Teams** roster with a **Match Satisfaction Report**.
- Closes the Round 2 form after matching, so late submissions don't arrive once teams are posted.

---

## One-time setup

### 1. Create the Round 1 Google Form
1. Create a new Google Form for students to vote on debate topics.
2. Add a **Multiple choice grid** question:
   - **Rows:** your candidate debate topics (e.g., AI Regulation, Universal Basic Income, Nuclear Energy).
   - **Columns:** ranks from `1` to `N`, where N is the number of topics and 1 is the top choice.
3. Turn on **Require a response in each row**.
4. Click the three dots (⋮) at the bottom right of the question and check **Limit to one response per column**.

> To change topics later, just edit the grid rows. Removed topics are ignored automatically.

### 2. Link the form to Google Sheets and add the script
1. In the Round 1 form, go to **Responses → Link to Sheets** and create a new spreadsheet.
2. Open the spreadsheet and go to **Extensions → Apps Script**.
3. Replace everything in `Code.gs` with the Debate Tools script, then save.
4. Refresh the spreadsheet. A **Debate Tools** menu appears.

The first time you run a menu item, Google will ask you to authorize the script to access your Forms and Sheets. Click **Review Permissions**, choose your account, and allow access.

The first time you run **1. Run Borda Count**, the script asks you to confirm which tab holds your Round 1 responses. Click the Round 1 responses tab first. The script remembers your answer.

---

## Running a debate cycle

### Step 0: Start a new cycle
Click **Debate Tools → 0. Start new debate cycle** *before* sharing the Round 1 form. This:
- Records the start time so Round 1 ignores ballots from earlier cycles.
- Clears last cycle's winners and Round 2 form, so the steps below start fresh.
- Reopens the Round 1 form if it was closed.

Old Borda Results and Final Teams tabs are kept.

### Step 1: Pick the two topics
Once students have voted, click **1. Run Borda Count (Round 1)**. A new **Borda Results** tab opens with the scores, the winners highlighted, and the number of ballots counted (plus how many older ballots were ignored).

### Step 2: Create the Round 2 form
Click **2. Generate Round 2 Form**. A popup shows the **student link** (to share) and the **edit link**. There are no manual settings to change on this form.

To get the links again later, use **❓ What do I do next?**

### Step 3: Match students into teams
After everyone submits Round 2, click **3. Match Teams (Round 2)**. A new **Final Teams** tab opens with:
- A **Match Satisfaction Report**: how many students got their 1st or 2nd topic choice and their preferred side.
- The full roster, grouped by topic and side, alphabetical within each group.

The Round 2 form closes automatically. If a student still needs to submit, reopen the form (**Responses → Accepting responses**) and run step 3 again. Each run creates a new Final Teams tab, so earlier rosters are kept.

---

## How the algorithms work

### Why a Borda count? (Round 1)
Instant runoff voting tends to favor topics with an enthusiastic plurality. A Borda count instead finds the topics with the broadest support across the whole class, which keeps down the number of students stuck debating a topic they strongly dislike.

With N topics, each ballot gives a topic:

$$\text{Points} = N - \text{rank} + 1$$

With 10 topics, a 1st-place ranking is worth 10 points, 2nd place 9 points, and 10th place 1 point.

### The 5-point preference gradient (Round 2)
**Topic assignment:** Students are shuffled randomly (Fisher-Yates) and placed in their first-choice topic until that topic is full. Everyone left over goes to the other topic.

**Side balancing:** Within each topic, each student's side preference becomes a score:

| Preference | Score |
|---|---|
| Strongly Prefer Pro | +2 |
| Prefer Pro | +1 |
| No Preference | 0 |
| Prefer Con | −1 |
| Strongly Prefer Con | −2 |

Students with the same score are put in random order. The list is sorted from +2 down to −2 and split in the middle: the top half is Pro, the bottom half Con. This means students with no preference or a mild one fill out team sizes before any student with a strong preference is moved to the other side.

---

## Troubleshooting

**"Click your Round 1 'Form Responses' tab and run this again."**
The script doesn't know your Round 1 tab yet. Click the tab linked to your Round 1 form and run step 1 again. You only have to do this once.

**"Couldn't find this cycle's Round 2 responses."**
No Round 2 form exists for this cycle yet (or it was deleted). Run step 2.

**"Missing Round 1 winners."**
Run step 1 before step 2.

**Tie warnings**
If topics tie for the second winning spot, the script picks alphabetically and names every tied topic in the alert. Check the Borda Results tab if you want to choose yourself. A tie for 1st place doesn't produce a warning, because both topics win either way.

**"The winners changed since your Round 2 form was created."**
You re-ran step 1 after making the Round 2 form, and the result changed. Run step 2 again and stop sharing the old Round 2 link.

**Students see "Use each rank only once."**
That's expected: they tried to give both topics the same rank.

---

## License
MIT
