import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QListWidget, QLabel, QTextEdit, QLineEdit, QComboBox, QVBoxLayout, QWidget
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from PyQt5.QtGui import QTextCursor, QFont

class ExcelEmailSender(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("E-Mails senden")
        self.setGeometry(100, 100, 800, 800)

        self.load_button = QPushButton("Excel-Datei öffnen", self)
        self.load_button.setGeometry(10, 10, 200, 30)
        self.load_button.clicked.connect(self.load_excel)

        self.email_list = QListWidget(self)
        self.email_list.setGeometry(10, 50, 580, 300)

        self.result_label = QLabel(self)
        self.result_label.setGeometry(10, 360, 580, 30)

        self.subject_label = QLabel("Betreff:", self)
        self.subject_label.setGeometry(10, 400, 100, 30)

        self.subject_input = QLineEdit(self)
        self.subject_input.setGeometry(120, 400, 470, 30)

        self.message_label = QLabel("Nachricht:", self)
        self.message_label.setGeometry(10, 440, 100, 30)

        self.message_input = QTextEdit(self)
        self.message_input.setGeometry(120, 440, 470, 100)

        self.bold_button = QPushButton("Fett", self)
        self.bold_button.setGeometry(600, 440, 80, 30)
        self.bold_button.clicked.connect(self.toggle_bold)

        self.italic_button = QPushButton("Kursiv", self)
        self.italic_button.setGeometry(690, 440, 80, 30)
        self.italic_button.clicked.connect(self.toggle_italic)

        self.underline_button = QPushButton("Unterstrichen", self)
        self.underline_button.setGeometry(600, 480, 170, 30)
        self.underline_button.clicked.connect(self.toggle_underline)

        self.font_size_label = QLabel("Schriftgröße:", self)
        self.font_size_label.setGeometry(600, 520, 100, 30)

        self.font_size_combo = QComboBox(self)
        self.font_size_combo.setGeometry(700, 520, 80, 30)
        self.font_size_combo.addItems(["8", "10", "12", "14", "16", "18", "20", "24"])
        self.font_size_combo.currentIndexChanged.connect(self.change_font_size)

        self.send_button = QPushButton("Senden", self)
        self.send_button.setGeometry(120, 560, 470, 40)
        self.send_button.clicked.connect(self.send_emails)

        self.sender_email = "mail"  # Hier die Absender-E-Mail-Adresse eintragen
        self.sender_password = "app-specific pw"  # Hier das Passwort für die Absender-E-Mail-Adresse eintragen

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Excel-Datei öffnen", "", "Excel-Dateien (*.xlsx *.xls)")
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.email_list.clear()
                for email in df['Email']:
                    self.email_list.addItem(email)
            except Exception as e:
                self.result_label.setText(f"Fehler beim Lesen der Excel-Datei: {str(e)}")

    def toggle_bold(self):
        cursor = self.message_input.textCursor()
        current_format = cursor.charFormat()
        current_format.setFontWeight(1 if current_format.fontWeight() == -1 else -1)
        cursor.mergeCharFormat(current_format)
        self.message_input.mergeCurrentCharFormat(current_format)

    def toggle_italic(self):
        cursor = self.message_input.textCursor()
        current_format = cursor.charFormat()
        current_format.setFontItalic(not current_format.fontItalic())
        cursor.mergeCharFormat(current_format)
        self.message_input.mergeCurrentCharFormat(current_format)

    def toggle_underline(self):
        cursor = self.message_input.textCursor()
        current_format = cursor.charFormat()
        current_format.setFontUnderline(not current_format.fontUnderline())
        cursor.mergeCharFormat(current_format)
        self.message_input.mergeCurrentCharFormat(current_format)

    def change_font_size(self):
        cursor = self.message_input.textCursor()
        current_format = cursor.charFormat()
        font_size = int(self.font_size_combo.currentText())
        current_format.setFontPointSize(font_size)
        cursor.mergeCharFormat(current_format)
        self.message_input.mergeCurrentCharFormat(current_format)

    def send_emails(self):
        subject = self.subject_input.text()
        message = self.message_input.toPlainText()

        if not subject or not message:
            self.result_label.setText("Betreff und Nachricht dürfen nicht leer sein.")
            return

        try:
            smtp_server = "smtp.gmail.com"  # Anpassen, wenn Sie einen anderen E-Mail-Anbieter verwenden
            smtp_port = 587  # Port für TLS-Verschlüsselung

            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(self.sender_email, self.sender_password)

            for row in range(self.email_list.count()):
                to_email = self.email_list.item(row).text()
                msg = MIMEMultipart()
                msg['From'] = self.sender_email
                msg['To'] = to_email
                msg['Subject'] = subject
                msg.attach(MIMEText(message, 'plain'))

                server.sendmail(self.sender_email, to_email, msg.as_string())

            server.quit()
            self.result_label.setText("E-Mails wurden erfolgreich gesendet.")
        except Exception as e:
            self.result_label.setText(f"Fehler beim Senden der E-Mails: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelEmailSender()
    window.show()
    sys.exit(app.exec_())
