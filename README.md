# JPI Exporter

Create a virtual environment.
``` Bash
python3 -m venv .venv
```

Activate the virtual environment.
``` Bash
source .venv/bin/activate
```

Install the requirements.
``` Bash
pip3 install requests openpyxl pyinstaller
```

Test the application.
``` Bash
python3 main.py
```

Create the distribution.
``` Bash
pyinstaller --onefile --name "JPI Exporter" main.py
```

Download via the GitHub link.
```
https://raw.githubusercontent.com/itopaloglu83/JPI-Exporter/main/Distribution/JPI%20Exporter.zip
```
