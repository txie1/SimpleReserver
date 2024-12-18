<a name="readme-top"></a>

# General-Purpose Reservation Sheets

This repo adds custom functionality to **Google Sheets** using Apps Scripts, to create a simple resource reservation system. It is designed for booking & managing shared resources such as equipment, conference rooms, or other assets.

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#features">Features</a>
    <li>
      <a href="#setup">Setup</a>
    <li>
      <a href="#usage">Example Usage</a>
    <li>
      <a href="#demo">Demo</a>
    <li>
      <a href="#guide">Customization Guide</a>
  <ol>
</details>


<a name="features"></a>
### Features
1. **Interface**: Reserve resources with a name, resource count, and purpose
2. **Visualization**: Updates reserved cells with reservation info
3. **Notifications**: Email all editors with summary of new reservation made
4. **Easy Resets**: Clear reservations for the previous day, or reset entire table

<img width="900" alt="example_panel" src="https://github.com/user-attachments/assets/4041404d-aac5-4bfb-b8a1-098b108f6497" />

----

<a name="setup"></a>
### Setup
1. Make a copy of this [Google Sheet](https://docs.google.com/spreadsheets/d/1A1IMIzoHbthuaOqvMi76wRDiE8JQ1SYESW3r6x81rQg/edit?usp=sharing)
2. Go to Extensions > **Apps Script**
3. Both scripts (*"Code.gs"*, *"ReservationalPanel.html"*) should already be copied, if not import them into editor
4. Select *"Code.gs"* and click `Run`
5. Click `advance` to allow permission to run scripts
6. Reload your Google Sheet – the `Reservation` menu will appear

<img width="500" alt="appscript" src="https://github.com/user-attachments/assets/e6ef82f8-e6d9-4791-bbc8-bb53a1abe9c5" />

<img width="450" alt="menu" src="https://github.com/user-attachments/assets/0c9b39bd-6884-4684-b29c-fc355da00f4f" />


<a name="usage"></a>
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
<a name="demo"></a>
### Demo
<img width="500" alt="confirm" src="https://github.com/user-attachments/assets/d674d710-6970-45c3-ac39-c0b47bfc3a38" />
<img width="500" alt="reserved cells" src="https://github.com/user-attachments/assets/825f4038-6727-4477-9322-44fd6745d822" />
<img width="300" alt="email" src="https://github.com/user-attachments/assets/011cb5da-02ad-4790-a32f-9ba81873cd5e" />


---
<a name="guide"></a>
### Customization Guide
Easily Modifiable

| **Category**          | **Details**                                                                                  |
|------------------------|---------------------------------------------------------------------------------------------|
| **Field Names**        | - Update sidebar labels and prompts (e.g., "Count", "Purpose", default color) in _`ReservationPanel.html`_  <br> - Replace terms like **"Resource"** with your specific asset name (e.g., "Room", "Equipment") in the script |
| **Email Content**      | - Modify the **subject line** and **body** in `sendEmailNotification()` to suit your use case  |
| **Time Slots** | - Adjust the reservation table size (e.g., **B2:H20**) in functions like `clearDailyReservations()` and `resetWeeklyReservations()` |


Things to Watch Out For

| **Category**          | **Details**                                                                                  |
|------------------------|---------------------------------------------------------------------------------------------|
| **Cell Mappings**      | - Ensure the header row (e.g., **Days of the Week** in row 1) aligns with the script logic.  <br> - Update column and row ranges in functions like `clearDailyReservations()` if your table structure changes |
| **Email Recipients**   | - The script sends notifications to all users with **edit access**.  <br> - Verify permissions if you need to limit email recipients |
| **Triggers**           | - Configure timed triggers carefully to avoid conflicts (e.g., overlapping with confirmation UI)  <br> - Check timezone settings: **File > Settings** in Google Sheets to ensure time-based triggers work as expected |



<p align="right">(<a href="#readme-top">back to top</a>)</p>
