# SimpleReserver
A script-enhanced Google Sheets solution for booking & managing shared resource reservations, implemented using `Apps Scripts`


### Features
1. **Interface**: Reserve resources with a name, resource count, and purpose
2. **Visualization**: Updates reserved cells with reservation info
3. **Notifications**: Email all editors with summary of new reservation made
4. **Easy Resets**: Clear reservations for the previous day, or reset entire table

<img width="900" alt="example_panel" src="https://github.com/user-attachments/assets/4041404d-aac5-4bfb-b8a1-098b108f6497" />

----

### Setup
1. Make a copy of this [Google Sheet](https://docs.google.com/spreadsheets/d/1A1IMIzoHbthuaOqvMi76wRDiE8JQ1SYESW3r6x81rQg/edit?usp=sharing)
2. Go to Extensions > **Apps Script**
3. Both scripts (*"Code.gs"*, *"ReservationalPanel.html"*) should already be copied, if not import them into editor
4. Select *"Code.gs"* and click `Run`
5. Click `advance` to allow permission to run scripts
6. Reload your Google Sheet – the `Reservation` menu will appear

<img width="500" alt="appscript" src="https://github.com/user-attachments/assets/e6ef82f8-e6d9-4791-bbc8-bb53a1abe9c5" />

<img width="450" alt="menu" src="https://github.com/user-attachments/assets/0c9b39bd-6884-4684-b29c-fc355da00f4f" />


### Example Usage
#### 1 - Reserve:

1. Open sidebar using `Reservation` -> `Open Reservation Panel`
2. Select cells (time slots), should be dynamically displayed on panel
3. Enter reservation info
4. Confirm reservation
5. Notification will be sent automatically after confirming

#### 2 - Sheet Updates:

- The cell(s) will be updated with user's name and resource count reserved
- Purpose/Notes will be added as comment to the cell(s)
- All editors will receive an email notification with a summary of reservation

#### 3 - Clearing Reservations:

- Use `Clear Yesterday` to remove reservations for the previous day
- Use `Clear All Entries` to clear the entire table
- (For automatic resets, add Triggers with steps below)


##### Steps to Add Timed Trigger
1. Open your Google Sheets and navigate to Extensions > Apps Script
2. In the Apps Script editor, go to `Triggers` section (clock icon in left menu)
3. Click on + Add Trigger (bottom-right corner)
4. Configure the trigger settings:
```
function: resetWeeklyReservations (or resetDailyReservations)
deployment: Head
event source: Time-driven
type of time based trigger: (e.g., Week timer > Every Monday)
```
5. Click Save. The function will be triggered and run at your preferred time

---

### Demo
<img width="600" alt="confirm" src="https://github.com/user-attachments/assets/d674d710-6970-45c3-ac39-c0b47bfc3a38" />
<img width="600" alt="reserved cells" src="https://github.com/user-attachments/assets/825f4038-6727-4477-9322-44fd6745d822" />
<img width="300" alt="email" src="https://github.com/user-attachments/assets/011cb5da-02ad-4790-a32f-9ba81873cd5e" />



