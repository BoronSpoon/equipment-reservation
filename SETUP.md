## Setting Up (administrator side) 
```
<details> 
<summary> 
(click to expand) 
</summary>
```

### 1. Create a Google account
### 2. Create a Google Group
<img align="right" display="block" width="40%" hspace="2%" src="pics/google_groups.gif">

1. Access https://groups.google.com/my-groups
2. Create group as shown in the right
   
<img width="100%" src="pics/horizontal_line_invisible.png">
3. Change settings as shown in below pictures

<img align="right" width="90%" src="pics/google_groups_settings.png">

<img width="100%" src="pics/horizontal_line_invisible.png">

### 3. Create a Google Apps Script Project 
1. Go to https://script.google.com/home and click New project
2. Copy the contents of [GitHub page](https://raw.githubusercontent.com/BoronSpoon/equipment-reservation/master/equipmentReservation.js) to the new project's source code

<img align="right" width="70%" src="pics/apps_script.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

3. In the following lines, 
   ```javascript
   properties.setProperty('groupUrl', '?????@googlegroups.com'); // groupURL
   properties.setProperty('timeZone', 'Asia/Tokyo'); // set timezone
   ```
   Change `?????@googlegroups.com` to the google group name created in the previous step.   
   Change `Asia/Tokyo` to your time zone. Refer to [wikipedia link](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones) for the syntax.

4. Add advanced calendar services and advanced sheets services as shown in following picture.
<img align="right" width="100%" src="pics/advanced_calendar_and_sheets.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

5. In the GAS project, run the function `setup()`.

<img align="right" width="55%" src="pics/apps_script_run_1.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

Click `Review Permissions`.

<img align="right" width="55%" src="pics/apps_script_run_2.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

Click `Advanced`.

<img align="right" width="55%" src="pics/apps_script_run_3.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

Click `Go to {project name} (unsafe)`.
   
<img align="right" width="55%" src="pics/apps_script_run_4.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

Click `Allow`. Don't touch any files and wait for **15** minutes.

<img align="right" width="55%" src="pics/apps_script_run_5.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

6. Deploy the project. Click `New Deployment`

<img align="right" width="25%" src="pics/apps_script_deploy_1.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

Click `Deploy`

<img align="right" width="55%" src="pics/apps_script_deploy_2.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

5. Add admin user
- Go to https://docs.google.com/spreadsheets and find a spreadsheet named `experimentConditionSpreadsheet`. Open the file and go to `users` sheet.   
- Change `ALL EVENTS` to something else as shown in below figure.

<img align="right" width="55%" src="pics/admin_user.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

6. Add equipments
- In `experimentConditionSpreadsheet`, go to `properties` sheet.   
- Add equipment names as shown in below figure.

<img align="right" width="35%" src="pics/equipment_name.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

- If possible, in each `equipment` sheet, enable filter as shown in below figure.
- this will auto sort reservations and hide canceled reservations.

<img align="right" width="60%" src="pics/enable_filter.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

7. Add equipment conditions
- In `experimentConditionSpreadsheet`, , go to `properties` sheet.   
- Add equipment conditions as shown in below figure.

<img align="right" width="35%" src="pics/experiment_condition.png">
<img width="100%" src="pics/horizontal_line_invisible.png">

</details>

<!-- USAGE EXAMPLES -->
## Setting Up (user side)
### 1. Join Google groups

<img align="right" display="block" width="35%" hspace="2%" src="pics/googlegroups1.png">   

1. Create a Google account
2. Access https://groups.google.com/forum/#!forum/{group_name}/join (ask the administrator for {group_name})
3. Select "join group"
<img width="100%" src="pics/horizontal_line_invisible.png">
<img align="left" display="block" width="30%" hspace="4%" src="pics/googlegroups2.png">   

4. Fill out the form like the following image. Use a display name that can be understood to the admin.
5. Subscription is complete. The security settings in the google groups is very strict, therefore you cannot see other members, and cannot post group emails.   

<img width="100%" src="pics/horizontal_line_invisible.png">

### 2. Initial settings
1. Adding you as an user
   1. Enter your name. Refer to ① in the below picture.
   2. Link for sheets：https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit?usp=sharing
2. Select devices you will use
   1. Check the devices you would like to use. Refer to ③ in the below picture.   

<img align="right" display="block" width="70%" hspace="4%" src="pics/sheets1.png">   
<img width="100%" src="pics/horizontal_line_invisible.png">

### 3. Add the user calendars to your calendar application (in your smartphone or on your PC)
1. Links for Read(read only) calendar and Write(write only) calendar are prepared. (Don't use the link in ➁)
2. Select the sheet named "URL of Calendar", and click the link like the figure below. Google calendar will open and you can add the calendar.   

<img align="right" display="block" width="70%" hspace="4%" src="pics/sheets2.png">   
<img width="100%" src="pics/horizontal_line_invisible.png">

### 4. Adding a reservation (Write calendar)
<img align="right" display="block" width="30%" src="pics/calendar1.png">   

1. Add an event on Google calendar
2. Title should be`[Device] [State]`. Refer to ① in the below picture. Refer to `naming conventions` for naming convention of Device and State.
3. Add the event to `Write [User Name]`. Refer to ① in the below picture.
<img width="100%" src="pics/horizontal_line_invisible.png">

### 5. naming conventions
- Devices (__You can change the name of devices in Sheets__)
-  | Device | Description |
   | ---- | ---- |
   | cmp  | Chemical mechanical planarization |
   | rie  | Reactive ion etching |
   | cvd  | Chemical vapor deposition |
   | pvd  | Physical vapor deposition |
   | euv  | Extreme ultraviolet lithography |
   | diff | Diffusion |
   | cleg | Cleaning |
   | dicg | Dicing |
   | pack | Packaging |
- States of equiment
-  | State | 説明 |
   | ---- | ---- |
   | evac | evacuation |
   | use/[NO ENTRY] | Main operation |
   | cool | cooling |
- Users (__You can add users in Sheets__)
-  | User Name | 名前 |
   | ---- | ---- |
   | Tana.Y1 | Tanaka Yuusuke |
   | Tana.Y2 | Tanabe Yuta |
   | Tana.S1 | Tanahashi Shion |
   | Chen.F1 | Chen Fan |
   | Zhu.Y1  | Zhu Yu |

### 6. Observing reservations (Write calendar)
1. Observe events on Google Calendar
2. Event title is `[Name] [Device] [State]`.
3. If there are too many reservations and it is overwhelming, you can hide the Read calendar in your calendar app and view it only when you need to.

### 7. Writing experiment conditions (Experiment Sheets)
1. Open Experiment Sheets
2. Find the current event that you created in Google Calendar
3. Write experiment conditions in Sheets
4. The conditions will automatically get updated to the event in Google \Calendar.

>The implementations details are in [details (Japanese)](https://github.com/rayanti/equipment-reservation/blob/master/details.md).

## Usage
- The equiment reservation infos are shown (as events) in Google Calendar. Clients will check the reservations on their PCs or phones, and make a reservation at an avaliable time (by creating events).
- Settings like display name, what equipments to show in Google Calendar are saved in a Google Sheet.
## About The Project
- This project uses a service called [Google Apps Script(GAS)](https://script.google.com/home/start). GAS enables Google services to be connected by scripts. We will use Google Sheets and Google Calendar in this project.
- The language used in GAS is JavaScript with V8 syntax (V8 is a runtime which is used in Chrome and Node.js).
- Sheets acts as a database for user information, preferences and logging. Calendar acts as an UI and is used for adding, modifying, reviewing event.
- There are two types of calendars we will use. The write calendar will be modified by the clients. The read calendar will be modified by the GAS service. When an event is modified on a write calendar, GAS applies the modification on the read calendar.

<!-- ROADMAP -->
## Roadmap
See the [open issues](https://github.com/rayanti/equipment-reservation/issues) for a list of proposed features (and known issues).

<!-- CONTRIBUTING -->
## Contributing
Contributions are what make the open source community such an amazing place to be learn, inspire, and create. Any contributions you make are **greatly appreciated**.
1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<!-- LICENSE -->
## License
Distributed under the MIT License. See `LICENSE` for more information.

<!-- ACKNOWLEDGEMENTS -->
## Acknowledgements
* [Best-README-Template](https://github.com/othneildrew/Best-README-Template)
