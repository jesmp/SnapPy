<h1>SnapPy</h1>

<p>SnapPy is a Python-based screenshot application designed to streamline the process of capturing screenshots while documenting steps in Gherkin language. The application allows users to take screenshots, annotate them, and generate a Word document with the captured images and corresponding steps. The application is primarily built for use in the Collections UAT department.</p>

<h2>Features:</h2>
<ul>
    <li>Capture Screenshots: Capture screenshots of your active window using Shift key, automatically save them, and insert them into a Word document.</li>
    <li>Gherkin Steps Integration: Input Gherkin language steps, with each screenshot associated with a specific step.</li>
    <li>Word Document Generation: Automatically generates a .docx file containing the screenshots and the associated Gherkin steps.</li>
    <li>Custom Save Path: Allows you to select a custom folder to save generated files.</li>
    <li>Notifications: Provides real-time notifications during the process for user actions and errors.</li>
    <li>Supports Multi-Monitor Setup: Works across all monitors to capture active window content.</li>
</ul>

<h2>Prerequisites:</h2>
<p>Before running the project, ensure you have the following installed:</p>
<ul>
    <li>Python 3.x: This project is built with Python 3.</li>
</ul>

<h2>Required Libraries:</h2>
<ul>
    <li>mss (for capturing screenshots)</li>
    <li>keyboard (for detecting keyboard input)</li>
    <li>pygetwindow (for getting active window details)</li>
    <li>pyinstaller (for packaging as an executable)</li>
    <li>plyer (for notifications)</li>
    <li>python-docx (for generating the Word document)</li>
</ul>

<p>To install the required libraries, you can use pip:</p>
<pre><code>pip install mss keyboard pygetwindow plyer python-docx</code></pre>

<h2>Installation</h2>
<ol>
    <li>
        Clone the repository:
        <pre><code>git clone https://github.com/your-username/SnapPy.git</code></pre>
    </li>
    <li>
        Install dependencies: Ensure that all required libraries are installed (see Prerequisites section).
    </li>
</ol>

<h2>Usage</h2>

<h3>1. Running the Application:</h3>
<p>To run the application, simply execute the following:</p>
<pre><code>python SnapPy.py</code></pre>
<p>This will launch the graphical user interface (GUI) of SnapPy where you can:</p>
<ul>
    <li>Set the document title.</li>
    <li>Optionally input an account/ECI number.</li>
    <li>Enter Gherkin steps for your documentation.</li>
    <li>Select the folder where files will be saved.</li>
    <li>Press Shift to capture a screenshot, Alt to move to the next step, and ` to check the current step.</li>
</ul>

<h3>2. Generating the Document:</h3>
<p>After completing the screenshots and steps, the application will generate a .docx file and save it to the selected folder with the title you provided. You will receive a notification once the file is saved.</p>

<h3>3. Save Path:</h3>
<p>The application will automatically remember the folder you selected for saving the files. You can update the save path at any time by clicking the Save Path button.</p>

<h2>Building the Executable</h2>

<p>If you want to package SnapPy as a standalone executable, you can use PyInstaller.</p>
<ol>
    <li>Install PyInstaller if you don’t have it:
    <pre><code>pip install pyinstaller</code></pre></li>
    <li>Run the following command to generate the .exe file:
    <pre><code>pyinstaller --onefile --noconsole --hidden-import="plyer.platforms.win.notification" --hidden-import="plyer.platforms" SnapPy.py</code></pre></li>
    <p>This will generate a .exe file in the dist folder.</p>
</ol>

<h2>License</h2>
<p>This project is licensed under the MIT License - see the LICENSE file for details.</p>

<h2>Author</h2>
<p>Jesus Peralta<br>Creator of SnapPy for Collections UAT.</p>

<h2>Troubleshooting</h2>

<h3>1. Common Issues:</h3>
<ul>
    <li><strong>Application Not Launching</strong>: If the application doesn't launch, ensure all the dependencies are installed properly. You can re-run <code>pip install -r requirements.txt</code> to ensure all libraries are installed.</li>
    <li><strong>Missing Libraries</strong>: If a required library is missing or not found, check the <code>requirements.txt</code> file for the exact versions and install them using pip.</li>
</ul>

<h3>2. Keyboard Shortcuts Not Working:</h3>
<ul>
    <li>Ensure that the <code>keyboard</code> library is correctly installed and that you have the necessary permissions to detect keyboard input, especially on Windows.</li>
</ul>

<h3>3. Word Document Not Generating:</h3>
<ul>
    <li>Verify that <code>python-docx</code> is installed. If you encounter issues, try uninstalling and reinstalling it with <code>pip install python-docx</code>.</li>
</ul>

<h3>4. Notifications Not Showing:</h3>
<ul>
    <li>Make sure <code>plyer</code> is correctly installed and working. If notifications aren’t appearing, check your operating system's notification settings.</li>
</ul>

<h2>Contributing</h2>
<p>Contributions are welcome! If you’d like to help improve SnapPy, feel free to fork the repository and submit a pull request. Please make sure to follow the project's code style and include tests for any new functionality.</p>

<h2>Contact</h2>
<p>If you have any questions, suggestions, or encounter issues, feel free to contact me via email at <a href="mailto:your.email@example.com">your.email@example.com</a>.</p>

</body>
</html>
