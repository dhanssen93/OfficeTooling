# Internal Exchange tool for simplifing dailey tasks

## Table of Contents
* [General Info](#general-information)
* [Technologies Used](#technologies-used)
* [Features](#features)
* [Setup](#setup)
* [Project Status](#project-status)
* [Wishlist](#Wishlist)
* [Room for Improvement](#room-for-improvement)
* [Contact](#contact)


## General Information
This personal project aims to simplify daily tasks in my organization. In the meantime I hope to gain more experience in the field of PowerShell scripting and tool building.


## Technologies Used
- PowerShell - version 5.1
- Exchange Online Management Shell - version 3.1.0
- ImportExcel - version 7.8.4


## Features
- Find out the address type of a specific email address. This could also be a distributionlist.
- Find out who is the owner of a specific mailbox, based on information in CustomAttribute1.
- Find out on which mailboxes a user has full access permission and of which distributionlistst the user is member of.
- Add a new owner to a mailbox documented in CustomAttribute1.
- Replace an old owner with a new one on the mailbox documented in CustomAttribute1.
- Audit who has rights to one or more mailboxes and export those results.
- Get an export of the mailboxsize of one or more mailboxes.


## Setup
What are the project requirements/dependencies? Where are they listed? A requirements.txt or a Pipfile.lock file perhaps? Where is it located?

Proceed to describe how to install / setup one's local environment / get started with the project.


## Project Status
Project is a work in progress.


## Wishlist
- Remove mailbox owner
- When checking for the owner, also list the owners where ; is replaced by , or "en"


## Room for Improvement
- Introduce try-catch where possible
- Add comment to the script
- Better menu for specific subjects


## Contact
Created by D. Hanssen > GitHub@visione.nl - feel free to contact me for any tips or suggestions!