import sys
import time
import socket
import os
import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QLineEdit, 
    QTextEdit, QFileDialog, QProgressBar, QMessageBox, QHBoxLayout, QGraphicsOpacityEffect
)
from PyQt5.QtCore import QThread, pyqtSignal, QPropertyAnimation, Qt

# ---------------------------
# Checkpoint Functions
# ---------------------------
CHECKPOINT_FILE = "progress_checkpoint.txt"

def save_checkpoint(index):
    """Save the current progress (row index) to a checkpoint file."""
    with open(CHECKPOINT_FILE, "w") as f:
        f.write(str(index))

def load_checkpoint():
    """Load the last processed row index from the checkpoint file."""
    if os.path.exists(CHECKPOINT_FILE):
        with open(CHECKPOINT_FILE, "r") as f:
            try:
                return int(f.read().strip())
            except:
                return 0
    return 0

# ---------------------------
# Utility Functions
# ---------------------------
def check_network():
    """
    Check for network connectivity by trying to connect to a public DNS.
    Returns True if network is available, otherwise False.
    """
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=5)
        return True
    except OSError:
        return False

def load_html_content(file_path, salutation, signature):
    """
    Loads the HTML file content and replaces placeholders if found.
    Otherwise, appends a signature.
    """
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()
        if '{sal}' in content or '{signature}' in content:
            content = content.replace("{sal}", salutation).replace("{signature}", signature)
        else:
            content += f"<br><br>Thanks & Regards,<br>{signature}"
        return content
    except Exception as e:
        return f"Error loading HTML file: {e}"

# ---------------------------
# Worker Thread Class
# ---------------------------
class EmailSenderWorker(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal()
    total_contacts_signal = pyqtSignal(int)
    emails_sent_signal = pyqtSignal(int)
    
    def __init__(self, excel_file, delay):
        super().__init__()
        self.excel_file = excel_file
        self.delay = delay
        self._pause = False
        self._running = True  
        self.failed_emails = []  # To store details of emails that failed to send.
    
    def pause(self):
        self._pause = True
        self.log_signal.emit("Process paused by user.")
    
    def resume(self):
        self._pause = False
        self.log_signal.emit("Process resumed by user.")
    
    def run(self):
        # Attempt to read the Excel file.
        try:
            df = pd.read_excel(self.excel_file)
        except Exception as e:
            self.log_signal.emit(f"Error reading Excel file: {e}")
            self.finished_signal.emit()
            return
        
        # Verify required columns.
        required_columns = ["from_email", "password", "sal", "signature", "to_email", "subject", "html_file"]
        for col in required_columns:
            if col not in df.columns:
                self.log_signal.emit(f"Missing required column: {col}")
                self.finished_signal.emit()
                return
        
        total = len(df)
        self.total_contacts_signal.emit(total)
        
        # Load checkpoint and resume from that row.
        start_index = load_checkpoint()
        self.log_signal.emit(f"Resuming from row {start_index}")
        
        sent_count = 0
        context = ssl.create_default_context()
        
        # Iterate over rows, starting from the last saved checkpoint.
        for i, row in df.iterrows():
            if i < start_index:
                continue  # Skip already processed rows
            
            # Respect manual pause if set.
            while self._pause:
                time.sleep(0.1)
            if not self._running:
                break
            
            sender = row["from_email"]
            password = row["password"]
            recipient = row["to_email"]
            subject = row["subject"]
            salutation = row["sal"]
            sender_name = row["signature"]  # Display name.
            html_file_path = row["html_file"]
            
            email_body = load_html_content(html_file_path, salutation, sender_name)
            
            # Build the email message.
            message = MIMEMultipart("alternative")
            message["Subject"] = subject
            message["From"] = f"{sender_name} <{sender}>"
            message["To"] = recipient
            message.attach(MIMEText(email_body, "html"))
            
            self.log_signal.emit(f"Sending email to {recipient} from {sender_name} (row {i})...")
            
            # Try sending the email with automatic network monitoring.
            email_sent = False
            while not email_sent:
                while self._pause:
                    time.sleep(0.1)
                try:
                    with smtplib.SMTP("smtp.gmail.com", 587, timeout=10) as server:
                        server.starttls(context=context)
                        server.login(sender, password)
                        server.sendmail(sender, recipient, message.as_string())
                    self.log_signal.emit("Email sent successfully!")
                    sent_count += 1
                    self.emails_sent_signal.emit(sent_count)
                    email_sent = True
                except Exception as e:
                    error_message = str(e)
                    # If network is down, pause sending until it's back.
                    if not check_network():
                        self.log_signal.emit("Network disconnected. Pausing email sending process. Waiting for network to be restored...")
                        while not check_network():
                            time.sleep(5)
                        self.log_signal.emit("Network restored. Resuming email sending process.")
                        # Automatically retry the same email.
                    else:
                        # For non-network errors, log and skip this email.
                        self.log_signal.emit(f"Failed to send email to {recipient}: {error_message}")
                        self.failed_emails.append({
                            'from_email': sender,
                            'to_email': recipient,
                            'subject': subject,
                            'error_message': error_message
                        })
                        break  # Exit the retry loop for this email.
            
            # Save checkpoint after processing each email.
            save_checkpoint(i + 1)
            self.progress_signal.emit(int(((i + 1) / total) * 100))
            self.log_signal.emit(f"Waiting for {self.delay} seconds before next email...\n")
            time.sleep(self.delay)
        
        # If there are failures, export a report.
        if self.failed_emails:
            try:
                df_failed = pd.DataFrame(self.failed_emails)
                report_filename = "failed_emails_report.xlsx"
                df_failed.to_excel(report_filename, index=False)
                self.log_signal.emit(f"Failure report saved to {report_filename}")
            except Exception as report_err:
                self.log_signal.emit(f"Error saving failure report: {report_err}")
        
        self.log_signal.emit("Finished sending emails.\n")
        self.finished_signal.emit()

# ---------------------------
# Main UI Class with Enhanced Attractive White Theme
# ---------------------------
class GmailSenderUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bulk Email Sender")
        self.resize(800, 700)
        self.excel_file = ""
        self.worker = None
        self.progress_anim = None  # To hold the progress bar animation.
        
        self.setup_ui()
        self.apply_styles()
        self.add_fade_in_animation()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Header label.
        self.header_label = QLabel("Bulk Email Sender")
        self.header_label.setAlignment(Qt.AlignCenter)
        self.header_label.setStyleSheet("font-size: 28pt; font-weight: bold;")
        layout.addWidget(self.header_label)
        
        # File selection layout.
        file_layout = QHBoxLayout()
        self.load_button = QPushButton("Load Excel File")
        self.load_button.clicked.connect(self.load_excel)
        file_layout.addWidget(self.load_button)
        
        self.file_label = QLabel("No file loaded")
        file_layout.addWidget(self.file_label)
        layout.addLayout(file_layout)
        
        # Delay input layout.
        delay_layout = QHBoxLayout()
        delay_label = QLabel("Delay between emails (seconds):")
        self.delay_input = QLineEdit("1")
        delay_layout.addWidget(delay_label)
        delay_layout.addWidget(self.delay_input)
        layout.addLayout(delay_layout)
        
        # Start sending button.
        self.send_button = QPushButton("Send Emails")
        self.send_button.clicked.connect(self.start_sending)
        layout.addWidget(self.send_button)
        
        # Pause and Resume buttons.
        pr_layout = QHBoxLayout()
        self.pause_button = QPushButton("Pause")
        self.pause_button.clicked.connect(self.pause_sending)
        self.pause_button.setEnabled(False)
        pr_layout.addWidget(self.pause_button)
        
        self.resume_button = QPushButton("Resume")
        self.resume_button.clicked.connect(self.resume_sending)
        self.resume_button.setEnabled(False)
        pr_layout.addWidget(self.resume_button)
        layout.addLayout(pr_layout)
        
        # Progress bar.
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # Display total contacts and emails sent.
        count_layout = QHBoxLayout()
        self.total_label = QLabel("Total Contacts: 0")
        self.sent_label = QLabel("Emails Sent: 0")
        count_layout.addWidget(self.total_label)
        count_layout.addWidget(self.sent_label)
        layout.addLayout(count_layout)
        
        # Log area.
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)
        
        self.setLayout(layout)
    
    def apply_styles(self):
        # Overall white theme with modern flat design.
        self.setStyleSheet("""
            QWidget {
                background-color: #FFFFFF;
                color: #333333;
                font-family: 'Segoe UI', sans-serif;
                font-size: 12pt;
            }
            QLabel {
                color: #333333;
            }
            QLineEdit, QTextEdit, QProgressBar {
                background-color: #F9F9F9;
                color: #333333;
                border: 1px solid #DDDDDD;
                border-radius: 5px;
                padding: 5px;
            }
            QProgressBar {
                border: 1px solid #DDDDDD;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #4CAF50, stop:1 #81C784);
                border-radius: 5px;
            }
        """)
        # Individual button styles with hover effects.
        self.load_button.setStyleSheet("""
            QPushButton {
                background-color:rgba(33, 243, 226, 0.69);
                color: white;
                padding: 10px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color:rgb(8, 95, 172);
            }
        """)
        self.send_button.setStyleSheet("""
            QPushButton {
                background-color:rgba(3, 124, 238, 0.68);
                color: white;
                padding: 10px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color:rgb(5, 22, 43);
            }
        """)
        self.pause_button.setStyleSheet("""
            QPushButton {
                background-color:rgba(255, 0, 221, 0.51);
                color: white;
                padding: 10px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color:rgba(91, 8, 226, 0.68);
            }
        """)
        self.resume_button.setStyleSheet("""
            QPushButton {
                background-color:rgb(16, 75, 114);
                color: white;
                padding: 10px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #8E24AA;
            }
        """)
    
    def add_fade_in_animation(self):
        """Adds a fade-in animation to the main window when it starts."""
        self.effect = QGraphicsOpacityEffect()
        self.setGraphicsEffect(self.effect)
        self.fade_anim = QPropertyAnimation(self.effect, b"opacity")
        self.fade_anim.setDuration(1000)  # Duration in milliseconds.
        self.fade_anim.setStartValue(0)
        self.fade_anim.setEndValue(1)
        self.fade_anim.start()
    
    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excel_file = file_path
            self.file_label.setText(file_path)
            self.log_area.append(f"Loaded file: {file_path}")
    
    def start_sending(self):
        if not self.excel_file:
            QMessageBox.critical(self, "Error", "Please load an Excel file first!")
            return
        try:
            delay = float(self.delay_input.text())
        except ValueError:
            QMessageBox.critical(self, "Error", "Please enter a valid number for delay!")
            return
        
        self.send_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.resume_button.setEnabled(False)
        
        self.worker = EmailSenderWorker(self.excel_file, delay)
        self.worker.log_signal.connect(self.update_log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.total_contacts_signal.connect(self.update_total)
        self.worker.emails_sent_signal.connect(self.update_sent)
        self.worker.start()
    
    def pause_sending(self):
        if self.worker is not None:
            self.worker.pause()
            self.pause_button.setEnabled(False)
            self.resume_button.setEnabled(True)
    
    def resume_sending(self):
        if self.worker is not None:
            self.worker.resume()
            self.pause_button.setEnabled(True)
            self.resume_button.setEnabled(False)
    
    def update_log(self, message):
        self.log_area.append(message)
    
    def update_progress(self, value):
        # Animate the progress bar value change.
        if self.progress_anim is not None:
            self.progress_anim.stop()
        self.progress_anim = QPropertyAnimation(self.progress_bar, b"value")
        self.progress_anim.setDuration(500)  # Animate over 0.5 seconds.
        self.progress_anim.setStartValue(self.progress_bar.value())
        self.progress_anim.setEndValue(value)
        self.progress_anim.start()
    
    def update_total(self, total):
        self.total_label.setText(f"Total Contacts: {total}")
    
    def update_sent(self, sent):
        self.sent_label.setText(f"Emails Sent: {sent}")
    
    def on_finished(self):
        self.send_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.resume_button.setEnabled(False)
        self.worker = None
        QMessageBox.information(self, "Finished", "Finished sending emails.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GmailSenderUI()
    window.show()
    sys.exit(app.exec_())
