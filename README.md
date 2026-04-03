<h1 align="center">📦 Get-Office</h1>

<p align="center">
  <b>Simple, automated Microsoft Office deployment & configuration script for Windows</b>
</p>

<p align="center">
  A lightweight batch script that <b>downloads, configures, installs, and activates</b> Microsoft Office — all from a single file.
</p>

<hr>

<h2>🖥️ Requirements</h2>
<ul>
  <li>Windows 10 / Windows 11</li>
  <li>Administrator privileges</li>
  <li>Internet connection</li>
</ul>

<hr>

<h2>⚡ Usage</h2>

<h3>📥 Method 1: Download</h3>
<ol>
  <li>Go to <a href="https://github.com/Reubak1/Get-Office/releases" target="_blank"><b>Releases</b></a></li>
  <li>Download <code>get-office.bat</code></li>
  <li>Run the script</li>
</ol>

<h3>⚡ Method 2: One-Line PowerShell</h3>
<p>Paste this into PowerShell:</p>

<pre><code>irm https://reubak1.github.io/Get-Office/ | iex</code></pre>

<p>⚠️ The script will automatically request administrator permissions if needed.</p>

<hr>

<h2>🔌 STEP 1: Architecture</h2>
<table>
  <tr>
    <th>Input</th>
    <th>Result</th>
  </tr>
  <tr>
    <td>Press <b>Y</b></td>
    <td>32-bit</td>
  </tr>
  <tr>
    <td>Press <b>N</b></td>
    <td>64-bit (recommended)</td>
  </tr>
</table>

<hr>

<h2>📱 STEP 2: App Selection</h2>
<p>Select which applications to install:</p>

<ul>
  <li>Access</li>
  <li>Excel</li>
  <li>Skype (Lync)</li>
  <li>OneDrive</li>
  <li>OneNote</li>
  <li>Outlook</li>
  <li>PowerPoint</li>
  <li>Publisher</li>
  <li>Word</li>
</ul>

<hr>

<h2>⬇️ STEP 3: Installation Process</h2>

<ul>
  <li>Generate configuration XML</li>
  <li>Download Office Deployment Tool</li>
  <li>Extract setup files</li>
  <li>Install Office</li>
  <li>Run activation</li>
  <li>Clean up temporary files</li>
</ul>

<hr>

<h2>⚙️ Configuration Details</h2>

<table>
  <tr><td><b>Edition</b></td><td>Office 2024 Volume License</td></tr>
  <tr><td><b>Channel</b></td><td>PerpetualVL2024</td></tr>
  <tr><td><b>Language</b></td><td>en-gb</td></tr>
  <tr><td><b>Updates</b></td><td>Enabled</td></tr>
</table>

<hr>

<h2>🧠 How It Works</h2>

<ul>
  <li><b>User Input → XML Generation</b></li>
  <li>Your choices determine which apps are excluded</li>
  <li>Office Deployment Tool is downloaded from a hosted source</li>
  <li>Silent install using <code>setup.exe /configure output.xml</code></li>
</ul>

<hr>

<h2>🔑 Activation</h2>

<p>
  Uses the <b>massgrave.dev activation script</b> hosted on this GitHub page.
  This method helps bypass ISP restrictions that may block the script.
</p>

<hr>

<h2>⚠️ Disclaimer</h2>

<p>
  This project is for <b>educational and testing purposes only</b>.
</p>

<ul>
  <li>Ensure you comply with Microsoft’s licensing terms</li>
  <li>Activation is provided by a third-party source</li>
</ul>

<hr>

<p align="center">
  ⭐ If you like this project, consider starring the repo!
</p>
