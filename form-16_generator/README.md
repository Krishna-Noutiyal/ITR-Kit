# Sola

## Overview 🕶️

Sola is a software tool designed to automate the Filling of Form-16 using ITR Format Excel file. It simplifies the process of creating Form-16 documents. The software extracts user data from their ITR-Format and Fills a pre-made Form-16 template. For using the software you will require the `ITR-Format` and `Form-16` templates which is not shared with the software.

> [!WARNING]
The software is created for **Pooja ITR Center**, and is publically available for use and contribution, but the Excel sheets like ITR-Format and Form-16 are kept private. Therefore, these sheets are not shared with this or any other software. **Without the Sheet Sola is a useless software !!**

## Features ✨

- Extracts details from ITR format ( `.xlsx` file ).
- Fills Form-16 ( `.xlsx` file ) automatically.
- User-friendly interface for file selection and output generation.
- Modern UI
- Can be compiled to web app, desktop app and apk.

## Installation ⬇️

To install the software you can download the latest relase of the software from the release section of this repository.

## Build ⚙️

To build the application from scratch for windows run the following command:

> [!NOTE]
> Make sure you have `git` and `tar` are installed on your system.
> To check if git and tar are present you can type `git` and `tar` directly to your console if no error is shown both are installed on your system.


```powershell
git archive --remote="https://github.com/Krishna-Noutiyal/ITR-Kit" HEAD:form-16_generator | tar -x
```
Head to the form-16_generator folder and run the build command:

```powershell
.\build.ps1 -i
```

The build script will automatically install the required packages and start the build. The finished build will be present in the `build\windows` directory.


## Usage ⚒️

For using Sola you will need ITR-Format and Form-16 template **( Not included with the software )**. But let's say you have the required documents then here is how you use it:

1. Launch the application.
2. Select the ITR Format.
3. Select the Form-16.
4. Click "Submit" to fill the Form-16.

> [!NOTE]
> The relevent details required to fill the **Form-16** is going to be automatically taken by Sola. Therefore, make sure you fill the **ITR-Format** before running Sola.

## File Structure 📂

The structure of the project is as follows :

- **assests/**: Contains the icon file for building Sola.
- **config/**: Color configurations of Sola.
- **icons/**: Software Icons used in Builds. The `assests\icon.png` file is the latest version of icon present in this dicrecory.
- **routes/**: Routes of the software and pages configuration.
- **scripts/**: Internal scripts used by sola for it's working.
- **ui/**: The Front-end user interface of the software.
- **README.md**: The thing you are reading now.
- **`./build.ps1`**: The build script of Sola. Also used to install reqirements.
- **`./main.py`**: The main file that starts the execution of Sola.
- **`./myproject.toml`**: Details of Sola project.
- **`./requirements.txt**: Packages required by Sola.

## Dependencies 🚴

- Python 3.9+
- Flet
- OpenPyXL
- Pandas
- Toml

## Sponsers and Funding 💰
The project is sponsered by **Pooja ITR Center**. All of the funding is provided by **Pooja ITR Center** with exclusive rights to the software.
