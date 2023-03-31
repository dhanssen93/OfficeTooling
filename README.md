# Internal Exchange tool for simplifing dailey tasks

## Table of Contents
* [General Info](#general-information)
* [Technologies Used](#technologies-used)
* [Features](#features)
* [Setup](#setup)
* [Project Status](#project-status)
* [Room for Improvement](#room-for-improvement)
* [Contact](#contact)


## General Information
This personal project aims to simplify daily tasks in my organization. In the meantime I hope to gain more experience in the field of PowerShell scripting and tool building.


## Technologies Used
- PowerShell - version 5.1
- Exchange Online Management Shell - version 3.1.0


## Features
List the ready features here:
- Find out the address type of a specific email address
- Find out who is the owner of a specific mailbox, based on information in CustomAttribute1
- Find out on which mailboxes a user has full access permission and of which distributionlistst the user is member of
- Audit who has rights to specific mailboxes
- Add a new owner to a mailbox documented in CustomAttribute1
- Replace an old owner with a new one on the mailbox documented in CustomAttribute1


## Setup
What are the project requirements/dependencies? Where are they listed? A requirements.txt or a Pipfile.lock file perhaps? Where is it located?

Proceed to describe how to install / setup one's local environment / get started with the project.


## Project Status
Project is a work in progress.


## Room for Improvement

Room for improvement:
- Introduce try-catch where possible
- Fix to remove ; at the beginnig of CustomAttribute1 when present
- Get-UserMailboxPermssions needs a check on if an address is filled in

To do:
- Add comment to the script
- Introduce new functions depending of the needs
- Fix bug where going to the menu doesn't work when using Get-MailboxPermissions


## Contact
Created by D. Hanssen > GitHub@visione.nl - feel free to contact me for any tips or suggestions!