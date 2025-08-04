#!/usr/bin/env python3
"""
Real-Time Screen Image Detection and Automation Tool
Single-file Python App for Windows 10/11

Requirements:
    pip install PySide6 opencv-python mss pyautogui win10toast pywin32 pillow pytesseract
"""

import sys
import time
import numpy as np
import cv2
import mss
import pyautogui
import winsound
import win32gui
from win10toast import ToastNotifier
import pytesseract
# If Tesseract-OCR is not in your PATH, uncomment and set the path below:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
from PIL import Image
from PySide6.QtCore import Qt, QRect, QPoint, Signal, QThread, QSize
from PySide6.QtGui import QPixmap, QIcon, QAction
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QListWidget, QListWidgetItem,
    QPushButton, QFileDialog, QSlider, QDoubleSpinBox, QCheckBox, QLineEdit,
    QHBoxLayout, QVBoxLayout, QFormLayout, QMessageBox, QSystemTrayIcon, QMenu,
    QRubberBand, QStyle
)

def bring_to_foreground(title_substring: str) -> None:
    def enum_callback(hwnd, results):
        title = win32gui.GetWindowText(hwnd).lower()
        if win32gui.IsWindowVisible(hwnd) and title_substring.lower() in title:
            results.append(hwnd)
    matches = []
    win32gui.EnumWindows(enum_callback, matches)
    if matches:
        hwnd = matches[0]
        win32gui.ShowWindow(hwnd, 5)
        win32gui.SetForegroundWindow(hwnd)

def grayscale(image: np.ndarray) -> np.ndarray:
    return cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

class ImageWatch:
    def __init__(self, path: str):
        self.path = path
        img = cv2.imread(path)
        self.template = grayscale(img) if img is not None else None
        if self.template is not None:
            self.h, self.w = self.template.shape[:2]
        else:
            self.h = self.w = 0
        self.threshold = 0.8
        self.region = None            # (x, y, w, h)
        self.move_mouse = False
        self.mouse_speed = 0.0
        self.click = False
        self.press_key = None
        self.notify = False
        self.sound = None
        self.window_title = None
        self.mask = None
        self.mask_path = None
        self.ocr_fallback = False
        self._triggered = False

class MonitorThread(QThread):
    error = Signal(str)
    def __init__(self, watch_list, interval: float = 0.1, ocr_lang: str = 'eng'):
        super().__init__()
        self.watch_list = watch_list
        self.interval = interval
        self.ocr_lang = ocr_lang
        self.toaster = ToastNotifier()
        self.running = False
    def run(self):
        with mss.mss() as sct:
            self.running = True
            while self.running:
                start = time.time()
                for item in list(self.watch_list):
                    try:
                        mon = (
                            {'top': item.region[1], 'left': item.region[0],
                             'width': item.region[2], 'height': item.region[3]}
                            if item.region else sct.monitors[1]
                        )
                        frame = np.array(sct.grab(mon))
                        gray = grayscale(frame[..., :3])
                        if item.template is None:
                            continue
                        if item.mask is not None:
                            res = cv2.matchTemplate(gray, item.template,
                                                    cv2.TM_CCOEFF_NORMED,
                                                    mask=item.mask)
                        else:
                            res = cv2.matchTemplate(gray, item.template,
                                                    cv2.TM_CCOEFF_NORMED)
                        _, max_val, _, max_loc = cv2.minMaxLoc(res)
                        found = max_val >= item.threshold
                        if not found and item.ocr_fallback:
                            text = pytesseract.image_to_string(
                                Image.fromarray(frame), lang=self.ocr_lang)
                            if 'skip' in text.lower():
                                found = True
                                max_loc = (mon['width']//2, mon['height']//2)
                        if found and not item._triggered:
                            item._triggered = True
                            x, y = max_loc
                            cx = mon.get('left', 0) + x + item.w // 2
                            cy = mon.get('top', 0) + y + item.h // 2
                            if item.move_mouse:
                                pyautogui.moveTo(cx, cy, duration=item.mouse_speed)
                            if item.click:
                                pyautogui.click()
                            if item.press_key:
                                pyautogui.press(item.press_key)
                            if item.notify:
                                self.toaster.show_toast(
                                    "Detected", f"{item.path} detected.", threaded=True
                                )
                            if item.sound:
                                winsound.PlaySound(
                                    item.sound,
                                    winsound.SND_FILENAME|winsound.SND_ASYNC
                                )
                            if item.window_title:
                                bring_to_foreground(item.window_title)
                        elif not found:
                            item._triggered = False
                    except Exception as e:
                        self.error.emit(str(e))
                elapsed = time.time() - start
                if elapsed < self.interval:
                    time.sleep(self.interval - elapsed)

class RegionSelector(QWidget):
    regionSelected = Signal(tuple)
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint|Qt.WindowStaysOnTopHint)
        self.setWindowOpacity(0.3)
        self.showFullScreen()
        self.origin = QPoint()
        self.rubber = QRubberBand(QRubberBand.Rectangle, self)
    def mousePressEvent(self, e):
        self.origin = e.pos()
        self.rubber.setGeometry(QRect(self.origin, QSize()))
        self.rubber.show()
    def mouseMoveEvent(self, e):
        self.rubber.setGeometry(QRect(self.origin, e.pos()).normalized())
    def mouseReleaseEvent(self, e):
        rect = self.rubber.geometry()
        self.rubber.hide()
        self.regionSelected.emit((rect.x(), rect.y(), rect.width(), rect.height()))
        self.close()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Screen Image Detection Tool")
        self.resize(900, 600)
        self.watch_list = []
        self.monitor_thread = None
        self._build_ui()
        self._build_tray()

    def _build_ui(self):
        container = QWidget()
        self.setCentralWidget(container)
        layout = QHBoxLayout(container)

        self.list_widget = QListWidget()
        self.list_widget.currentItemChanged.connect(self._on_select)
        layout.addWidget(self.list_widget, 1)

        form = QFormLayout()
        self.preview_label = QLabel("No image selected")
        self.preview_label.setFixedSize(200, 200)
        self.preview_label.setAlignment(Qt.AlignCenter)
        form.addRow(self.preview_label)

        self.sensitivity_slider = QSlider(Qt.Horizontal)
        self.sensitivity_slider.setRange(1, 100)
        self.sensitivity_slider.valueChanged.connect(self._on_sensitivity)
        form.addRow("Sensitivity (%)", self.sensitivity_slider)

        rl = QHBoxLayout()
        self.region_edit = QLineEdit()
        self.region_edit.setPlaceholderText("x,y,width,height")
        self.region_edit.editingFinished.connect(self._on_region_text)
        btn_region = QPushButton("Select Region")
        btn_region.clicked.connect(self._select_region)
        rl.addWidget(self.region_edit)
        rl.addWidget(btn_region)
        form.addRow("Region", rl)

        self.move_checkbox = QCheckBox()
        self.move_checkbox.stateChanged.connect(self._on_move_toggle)
        form.addRow("Move Mouse", self.move_checkbox)

        self.speed_spinbox = QDoubleSpinBox()
        self.speed_spinbox.setRange(0.0, 5.0)
        self.speed_spinbox.valueChanged.connect(self._on_speed_change)
        form.addRow("Speed (s)", self.speed_spinbox)

        self.click_checkbox = QCheckBox()
        self.click_checkbox.stateChanged.connect(self._on_click_toggle)
        form.addRow("Click", self.click_checkbox)

        self.key_edit = QLineEdit()
        self.key_edit.editingFinished.connect(self._on_key_press)
        form.addRow("Key Press", self.key_edit)

        self.notify_checkbox = QCheckBox()
        self.notify_checkbox.stateChanged.connect(self._on_notify_toggle)
        form.addRow("Windows Toast", self.notify_checkbox)

        sl = QHBoxLayout()
        self.sound_edit = QLineEdit()
        btn_sound = QPushButton("Browse")
        btn_sound.clicked.connect(self._browse_sound)
        sl.addWidget(self.sound_edit)
        sl.addWidget(btn_sound)
        form.addRow("Sound File", sl)

        self.ocr_checkbox = QCheckBox()
        self.ocr_checkbox.stateChanged.connect(self._on_ocr_toggle)
        form.addRow("Use OCR Fallback", self.ocr_checkbox)

        ml = QHBoxLayout()
        self.mask_edit = QLineEdit()
        btn_mask = QPushButton("Load Mask")
        btn_mask.clicked.connect(self._browse_mask)
        ml.addWidget(self.mask_edit)
        ml.addWidget(btn_mask)
        form.addRow("Template Mask", ml)

        self.window_edit = QLineEdit()
        self.window_edit.editingFinished.connect(self._on_window_activate)
        form.addRow("Activate Window", self.window_edit)

        bl = QHBoxLayout()
        btn_add = QPushButton("Add Image")
        btn_add.clicked.connect(self._on_add)
        btn_remove = QPushButton("Remove Image")
        btn_remove.clicked.connect(self._on_remove)
        bl.addWidget(btn_add)
        bl.addWidget(btn_remove)
        form.addRow(bl)

        self.start_button = QPushButton("Start Monitoring")
        self.start_button.clicked.connect(self._on_toggle)
        form.addRow(self.start_button)

        right_widget = QWidget()
        right_widget.setLayout(form)
        layout.addWidget(right_widget, 2)

    def _build_tray(self):
        icon = QIcon.fromTheme("applications-system")
        self.tray = QSystemTrayIcon(icon, self)
        menu = QMenu()
        show_act = QAction("Show", self)
        show_act.triggered.connect(self.show)
        exit_act = QAction("Exit", self)
        exit_act.triggered.connect(self._on_exit)
        menu.addAction(show_act)
        menu.addAction(exit_act)
        self.tray.setContextMenu(menu)
        self.tray.show()

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.tray.showMessage(
            "Running in background", "App is still running.", QSystemTrayIcon.Information, 2000
        )

    def _on_exit(self):
        if self.monitor_thread and self.monitor_thread.isRunning():
            self.monitor_thread.running = False
            self.monitor_thread.wait()
        QApplication.quit()

    def _on_add(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Image", "", "Images (*.png *.jpg)")
        for p in paths:
            w = ImageWatch(p)
            self.watch_list.append(w)
            it = QListWidgetItem(p.split('/')[-1])
            it.setData(Qt.UserRole, w)
            self.list_widget.addItem(it)

    def _on_remove(self):
        row = self.list_widget.currentRow()
        if row >= 0:
            self.watch_list.pop(row)
            self.list_widget.takeItem(row)

    def _on_select(self, current, previous=None):
        if current:
            w = current.data(Qt.UserRole)
            pix = QPixmap(w.path).scaled(self.preview_label.size(), Qt.KeepAspectRatio)
            self.preview_label.setPixmap(pix)
            self.sensitivity_slider.setValue(int(w.threshold * 100))
            self.region_edit.setText(','.join(map(str, w.region)) if w.region else '')
            self.move_checkbox.setChecked(w.move_mouse)
            self.speed_spinbox.setValue(w.mouse_speed)
            self.click_checkbox.setChecked(w.click)
            self.key_edit.setText(w.press_key or '')
            self.notify_checkbox.setChecked(w.notify)
            self.sound_edit.setText(w.sound or '')
            self.ocr_checkbox.setChecked(w.ocr_fallback)
            self.mask_edit.setText(w.mask_path or '')
            self.window_edit.setText(w.window_title or '') 

    def _on_sensitivity(self, val):
        item = self.list_widget.currentItem()
        if item:
            item.data(Qt.UserRole).threshold = val / 100.0

    def _on_region_text(self):
        item = self.list_widget.currentItem()
        if item:
            try:
                parts = self.region_edit.text().split(',')
                coords = list(map(int, parts))
                if len(coords) == 4:
                    item.data(Qt.UserRole).region = tuple(coords)
            except:
                pass

    def _select_region(self):
        selector = RegionSelector()
        selector.regionSelected.connect(self._set_region)
        selector.show()

    def _set_region(self, region):
        item = self.list_widget.currentItem()
        if item:
            item.data(Qt.UserRole).region = region
            self.region_edit.setText(','.join(map(str, region)))

    def _on_move_toggle(self, state):
        item = self.list_widget.currentItem()
        if item:
            item.data(Qt.UserRole).move_mouse = bool(state)

    def _on_speed_change(self, value):
        item = self.list_widget.currentItem()
        if item:
            item.data(Qt.UserRole).mouse_speed = value

    def _on_click_toggle(self, state):
        item = self.list_widget.currentItem()
        if item:
            item.data(Qt.UserRole).click = bool(state)

    def _on_key_press(self):
        item = self.list_widget.currentItem()
        if item:
            item.data(Qt.UserRole).press_key = self.key_edit.text() or None

    def _on_notify_toggle(self, state):
        item = self.list_widget.currentItem()
        if item:
            item.data(Qt.UserRole).notify = bool(state)

    def _browse_sound(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Sound", "", "WAV Files (*.wav)")
        if path:
            item = self.list_widget.currentItem()
            if item:
                item.data(Qt.UserRole).sound = path
                self.sound_edit.setText(path)

    def _on_ocr_toggle(self, state):
        item = self.list_widget.currentItem()
        if item:
            item.data(Qt.UserRole).ocr_fallback = bool(state)

    def _browse_mask(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Mask", "", "Images (*.png *.jpg)")
        if path:
            img_mask = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
            _, mask = cv2.threshold(img_mask, 1, 255, cv2.THRESH_BINARY)
            item = self.list_widget.currentItem()
            if item:
                w = item.data(Qt.UserRole)
                w.mask = mask
                w.mask_path = path
                self.mask_edit.setText(path)

    def _on_window_activate(self):
        item = self.list_widget.currentItem()
        if item:
            item.data(Qt.UserRole).window_title = self.window_edit.text() or None

    def _on_toggle(self):
        if self.monitor_thread and self.monitor_thread.isRunning():
            self.monitor_thread.running = False
            self.monitor_thread.wait()
            self.start_button.setText("Start Monitoring")
        else:
            self.monitor_thread = MonitorThread(self.watch_list)
            self.monitor_thread.error.connect(lambda msg: QMessageBox.critical(self, "Error", msg))
            self.monitor_thread.start()
            self.start_button.setText("Stop Monitoring")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
