# XML-Schnittstelle

XML Generation from Excel Data.  ----- German Version Below ------
Description

This Python script converts data from an Excel file into a structured XML format suitable for various applications requiring XML data input. It is designed to handle specific data fields and convert them according to pre-defined XML schema requirements.
Prerequisites

Before running this script, ensure you have Python installed on your system. This script is compatible with Python 3.x.

    Download Python
    Libraries used:
        pandas
        datetime
        xml.etree.ElementTree
        re
        pytz

Installation

    Clone the Repository:

    bash

git clone https://github.com/yourusername/your-repository-name.git
cd your-repository-name

Install Required Python Libraries:

bash

    pip install pandas pytz

    Prepare the Environment:
        Ensure you have an Excel file named DELA-EXCEL6k.xlsx in the same directory as the script. If your file has a different name or location, adjust the file_path variable in the script accordingly.

Configuration

    Excel File Setup:
        Place your Excel file (DELA-EXCEL6k.xlsx) in the project directory.
        If your file is named differently, open the script and modify the file_path variable to match your file's name:

        python

        file_path = 'YourExcelFileName.xlsx'

Running the Script

    Execute the Script:

    open a terminal in the directory where the script and the excel list are

    python3 gen2.py

    

    Output:
        The script will generate an XML file named Delafinal6k.xml in the project directory.

Post-Processing

After generating the XML file:

    Replace the First Line: Open the generated XML file and replace the first line with the contents of the header file (if applicable).
    (important !!) Add Closing Tag: Ensure the </OPENQCAT> tag is present at the end of the XML file to properly close the XML structure.



-------------------------German --------------------
Projekt-Titel

XML-Schnittstelle aus Excel-Daten
Beschreibung

Dieses Python-Skript konvertiert Daten aus einer Excel-Datei in ein strukturiertes XML-Format, das für verschiedene Anwendungen geeignet ist, die XML-Dateneingaben erfordern. Es ist darauf ausgelegt, spezifische Datenfelder zu verarbeiten und entsprechend vordefinierten XML-Schema-Anforderungen zu konvertieren.
Voraussetzungen

Bevor Sie dieses Skript ausführen, stellen Sie sicher, dass Python auf Ihrem System installiert ist. Dieses Skript ist kompatibel mit Python 3.x.

    Python herunterladen
    Verwendete Bibliotheken:
        pandas
        datetime
        xml.etree.ElementTree
        re
        pytz

Installation

    Repository klonen:

    im terminal 

git clone https://github.com/IhrBenutzername/IhrRepositoryName.git
cd IhrRepositoryName

Erforderliche Python-Bibliotheken installieren:

im terminal

    pip install pandas pytz

    Umgebung vorbereiten:
        Stellen Sie sicher, dass Sie eine Excel-Datei mit dem Namen DELA-EXCEL6k.xlsx im gleichen Verzeichnis wie das Skript haben. Wenn Ihre Datei anders heißt oder sich an einem anderen Ort befindet, passen Sie die Variable file_path im Skript entsprechend an.

Konfiguration

    Excel-Datei einrichten:
        Platzieren Sie Ihre Excel-Datei (DELA-EXCEL6k.xlsx) im Projektverzeichnis.
        Wenn Ihre Datei anders benannt ist, öffnen Sie das Skript und ändern Sie die Variable file_path, um sie an den Namen Ihrer Datei anzupassen:

        python

        file_path = 'IhrExcelDateiName.xlsx'

Das Skript ausführen

    Skript ausführen:

    bash

    python gen2.py

    

    Ausgabe:
        Das Skript generiert eine XML-Datei mit dem Namen Delafinal6k.xml im Projektverzeichnis.

Nachbearbeitung

Nach der Generierung der XML-Datei:

    Erste Zeile ersetzen: Öffnen Sie die generierte XML-Datei und ersetzen Sie die erste Zeile durch den Inhalt der Header-Datei (falls zutreffend).
    (Wichtig !!!)Abschließenden Tag hinzufügen: Stellen Sie sicher, dass der Tag </OPENQCAT> am Ende der XML-Datei vorhanden ist, um die XML-Struktur ordnungsgemäß zu schließen.



