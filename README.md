SnapPy

SnapPy is a Python-based screenshot application designed to streamline the process of capturing screenshots while documenting steps in Gherkin language. The application allows users to take screenshots, annotate them, and generate a Word document with the captured images and corresponding steps. The application is primarily built for use in the Collections UAT department.

Features
	•	Capture Screenshots:
Capture screenshots of your active window using Shift key, automatically save them, and insert them into a Word document.
	•	Gherkin Steps Integration:
Input Gherkin language steps, with each screenshot associated with a specific step.
	•	Word Document Generation:
Automatically generates a .docx file containing the screenshots and the associated Gherkin steps.
	•	Custom Save Path:
Allows you to select a custom folder to save generated files.
	•	Notifications:
Provides real-time notifications during the process for user actions and errors.
	•	Supports Multi-Monitor Setup:
Works across all monitors to capture active window content.

Prerequisites

Before running the project, ensure you have the following installed:
	•	Python 3.x: This project is built with Python 3.
	•	Required Libraries:
	•	mss (for capturing screenshots)
	•	keyboard (for detecting keyboard input)
	•	pygetwindow (for getting active window details)
	•	pyinstaller (for packaging as an executable)
	•	plyer (for notifications)
	•	python-docx (for generating the Word document)

To install the required libraries, you can use pip:

pip install mss keyboard pygetwindow plyer python-docx

Installation
	1.	Clone the repository:

git clone https://github.com/your-username/SnapPy.git

	2.	Install dependencies:
Ensure that all required libraries are installed (see Prerequisites section).

Usage
	1.	Running the Application:
To run the application, simply execute the following:

python SnapPy.py

This will launch the graphical user interface (GUI) of SnapPy where you can:
	•	Set the document title.
	•	Optionally input an account/ECI number.
	•	Enter Gherkin steps for your documentation.
	•	Select the folder where files will be saved.
	•	Press Shift to capture a screenshot, Alt to move to the next step, and ` to check the current step.

	2.	Generating the Document:
After completing the screenshots and steps, the application will generate a .docx file and save it to the selected folder with the title you provided. You will receive a notification once the file is saved.
	3.	Save Path:
The application will automatically remember the folder you selected for saving the files. You can update the save path at any time by clicking the Save Path button.

Building the Executable

If you want to package SnapPy as a standalone executable, you can use PyInstaller.
	1.	Install PyInstaller if you don’t have it:

pip install pyinstaller

	2.	Run the following command to generate the .exe file:

pyinstaller --onefile --noconsole --hidden-import="plyer.platforms.win.notification" --hidden-import="plyer.platforms" SnapPy.py

This will generate a .exe file in the dist folder.

License

This project is licensed under the MIT License - see the LICENSE file for details.

Author

Jesus Peralta
Creator of SnapPy for Collections UAT.

This README.md provides an overview of the SnapPy application, setup instructions, and usage details. Let me know if you need any changes or additions!