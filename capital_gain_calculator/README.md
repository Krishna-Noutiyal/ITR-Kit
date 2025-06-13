# CGC

## Overview 🕶️

CGC is a capital gain vaisualizer designed to automate the calculations of **Stock Market Trades**. It simplifies the process of tax calculations on capital gains by **Stock Markte Trades** . The software extracts trade data from **AIS's Trade Data** and creates `Capital Gain.xlsx` for visualization of **Stock Market Trades** and tax calculation on thoes trades. For using the software you will require the **`AIS's Trade Data`** which can be downloaded from the **[E-Filling Portal](https://eportal.incometax.gov.in/)**.

> [!WARNING]
The software is created for **Pooja ITR Center**, and is publically available for use and contribution, but the Excel sheets like `ITR-Format` and `Form-16` are kept private. Therefore, these sheets are not shared with this or any other software. **CGC doesn't need `ITR-Format` or `Form-16` to work !!**

## Features ✨

- Extracts trade details from multiple **Trade CSVs** downloaded from **[E-Filling Portal](https://eportal.incometax.gov.in/)**
- Generated `Capital Gain.xlxs` for visualizing Tax and Trade Calculationi.
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
git archive --remote="https://github.com/Krishna-Noutiyal/ITR-Kit" HEAD:capital_gain_calculator | tar -x
```
Head to the capital_gain_calculator folder and run the build command:

```powershell
.\build.ps1 -i
```

The build script will automatically install the required packages and start the build. The finished build will be present in the `build\windows` directory.


## Usage ⚒️

For using CGC you will need Trading Data, the Trading data can be downloaded from **[E-Filling Portal](https://eportal.incometax.gov.in/)**, the data will be in multiple CSVs that contains all of the short and long term trades performed over the current Financial Year. After downloading the CSVs follow these steps:

1. Launch the application.
2. Select the **Trade CSVs**.
3. Select the Output path.
4. Click "Submit" to Create stunning `Capital Gain.xlsx` dashboard.

## File Structure 📂

The structure of the project is as follows :

- **Dashboards** : Dashboads used to visualize **Trade Data**.
- **assests** : Contains the icon file for building CGC.
- **config** : Color configurations of CGC.
- **icons** : Software Icons used in Builds. The `assests\icon.png` file is the latest version of icon present in this dicrecory.
- **routes** : Routes of the software and pages configuration.
- **scripts** : Internal scripts used by CGC for it's working.
- **ui** : The Front-end user interface of the software.
- **README.md** : The thing you are reading now.
- **`./build.ps1`** : The build script of CGC. Also used to install reqirements.
- **`./main.py`** : The main file that starts the execution of CGC.
- **`./myproject.toml`** : Details of CGC project.
- **`./requirements.txt** : Packages required by CGC.

## Dependencies 🚴

- Python 3.9+
- Flet
- Xlsxwriter
- Pandas
- Toml

## Sponsers and Funding 💰
The project is sponsered by **Pooja ITR Center**. All of the funding is provided by **Pooja ITR Center** with exclusive rights to the software.
