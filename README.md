📦 Get-Office 

Simple, automated Microsoft Office deployment and configuration script for Windows.
Get-Office is a lightweight batch script that downloads, configures, installs, and activates Microsoft Office with user-selected components from a single batch file.

🖥️ Requirements

Windows 10 / Windows 11
Administrator privileges
Internet connection

⚡ Usage

Go to releases and select get-office.bat


⚠️ The script will automatically request admin permissions if needed


🔌 STEP 1: Architecture

Press Y → 32-bit

Press N → 64-bit (recommended)


📱 STEP 2: App Selection

Choose which apps to install:

Access

Excel

Skype (Lync)

OneDrive

OneNote

Outlook

PowerPoint

Publisher

Word

⬇️ STEP 3: Installation Process

The script will:

Generate a configuration XML


Download the Office Deployment Tool

Extract setup files

Install Office

Run activation

Clean up temporary files

⚙️ Configuration Details

Edition: Office 2024 Volume License


Channel: PerpetualVL2024

Language: en-gb

Updates: Enabled

🧠 How It Works

User Input → XML Generation

Your choices determine which apps are excluded

Office Deployment Tool downloaded directly from a hosted source

Silent Installation

Uses setup.exe /configure output.xml

🔑 Activation

Uses massgrave.dev activation script hosted on this github page to bypass internet service providers blocking the script

⚠️ Disclaimer

This project is for educational and testing purposes only.

Ensure you comply with Microsoft’s licensing terms.

The activation method is provided by a third-party source.
