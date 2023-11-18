# Outlook Delete Empty Inbox Folders

![App Icon](/icon/folder-128.png)

A Windows application (written in Delphi (Object Pascal) that deletes all empty Inbox folders in the default Outlook profile on the system running the app. 

[Icon File](https://www.flaticon.com/free-icon/folder_4080842?term=folder+delete&page=1&position=31&origin=search&related_id=4080842); Icon Attribution: <a href="https://www.flaticon.com/free-icons/delete" title="delete icons">Delete icons created by Freepik - Flaticon</a>

## Tasks

- [ ] Add About page
- [ ] Publish to Windows store

## About the Application

Outlook users that leverage Outlook's Archive feature to move older items to an archive file eventually end up creating empty folders in their Inbox when Outlook archives all of the mail messages to the archive. Users can manually delete them, but that process is time consuming and tedius if you have a lof of folders. 

This application is a simple utility for Microsoft Windows that automatically deletes all empty Inbox folders in a user's default Outlook Profile. 

When you start the application, a splash screen appears for two seconds as shown below.

![Application Splash Screen](/images/figure-01.png)

When the main screen appears, click the **Analyze** button to scan the default profile's Inbox for empty folders. 

![Application Main Screen](/images/figure-02.png)

When the application finds an empty Inbox folder, even in a sub-folder, it adds the folder path to the sorted list of folders on the application's main screen. The application displays the number of empty folders in the status bar at the bottom of the application window.

No folder deletion happens yet, you must first review the list and validate that these are the correct folders to delete (even checking in Outlook to make sure). When you are confident that the list is correct, click the **Delete** button to start the process of deleting all of the empty folders. 

**Note:** Folder deletion is not reversable, once the application deletes the folder, there's no possible recovery of the deleted folder. Please make a backup copy of your Outlook PST file before running the application.

For troubleshooting purposes, the application creates a log file every time the application executes. The file is called `DeleteEmptyFolders.log` and the application creates the file in the default Windows `temp` folder. To locate the system's `temp` folder, open Windows Explorer and enter `%temp%` in the location field at the top of the Explorer window as shown below.

![Windows Explorer](/images/figure-03.png)

When you press the keyboard's **Enter** key, Explorer will open the system's `temp` folder as shown in the following figure.

![Windows Explorer Temp Folder](/images/figure-04.png)

The contents of the file look something like the following text.

```
11/18/2023 11:23:21 AM Log opened
11/18/2023 11:23:21 AM Form activated
11/18/2023 11:23:25 AM Analyze button clicked
11/18/2023 11:23:25 AM Loading Outlook folders
11/18/2023 11:23:25 AM Delete disabled
11/18/2023 11:23:27 AM Folder: Inbox\ (13)
11/18/2023 11:23:27 AM Folder: Inbox\Events\
11/18/2023 11:23:27 AM Folder: Inbox\Product Support\
11/18/2023 11:23:27 AM Folder: Inbox\Product Support\ (7)
11/18/2023 11:23:27 AM Folder: Inbox\Product Support\Sonos\
11/18/2023 11:23:27 AM Deletable: Inbox\Product Support\Sonos\
11/18/2023 11:23:27 AM Folder: Inbox\Product Support\Dymo\
11/18/2023 11:23:27 AM Deletable: Inbox\Product Support\Dymo\
11/18/2023 11:23:27 AM Folder: Inbox\Product Support\Particle\
11/18/2023 11:23:27 AM Deletable: Inbox\Product Support\Particle\
11/18/2023 11:23:27 AM Folder: Inbox\Travel\
11/18/2023 11:23:27 AM Folder: Inbox\Concerts\
11/18/2023 11:23:27 AM Folder: Inbox\Product Orders\
11/18/2023 11:23:27 AM Folder: Inbox\Product Orders\ (42)
11/18/2023 11:23:27 AM Folder: Inbox\Product Orders\Sticker Mule\
11/18/2023 11:23:27 AM Folder: Inbox\Product Orders\Home Depot\
11/18/2023 11:23:27 AM Folder: Inbox\Product Orders\Sonos\
11/18/2023 11:23:27 AM Folder: Inbox\Product Orders\Etsy\
11/18/2023 11:23:27 AM Folder: Inbox\Product Orders\Music\
.
.
.
11/18/2023 11:23:37 AM View Log button clicked
11/18/2023 11:44:21 AM Closing application
```

Looking through the file, you can see every Inbox folder the application identified as well as any it thinks is deletable. The number in parenthesis after the folder path indicates the number of sub-folders in the current folder path.


## Limitations

The application only works with the default Outlook profile on the system. It's likely possible to build in the ability to let the user select the profile to use, but the application today does not do that. This project is open source, so if you want to make that change, please save the changes back so other users can use that feature,

The application was only tested against a local PST file (not an Offline Outlook Data file (`.ost`) from Microsoft Exchange). It may work with an Offline Outlook Data file, I truly have no idea.


## Build Requirements 

The application uses several [Konopka Signature VCL Controls](https://www.componentsource.com/product/konopka-signature-vcl-controls) formerly known as Raize Components. They ship free with Delphi now, so you will just have to install the package using the GetIt Package Manager to build this app on your own.

***

If this code helps you, please consider buying me a coffee.

<a href="https://www.buymeacoffee.com/johnwargo" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" height="41" width="174"></a>