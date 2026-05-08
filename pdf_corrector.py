import sys
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QListWidget, QGraphicsView, QGraphicsScene, 
                             QGraphicsPixmapItem, QSlider, QLabel, QPushButton, 
                             QComboBox, QButtonGroup, QSplitter, QDoubleSpinBox, 
                             QFileDialog, QListWidgetItem, QCheckBox, QMessageBox, 
                             QGraphicsRectItem, QGraphicsObject, QGraphicsItem, QListView)
from PyQt6.QtGui import QPixmap, QPainter, QAction, QImage, QIcon, QTransform, QPen, QColor, QBrush, QCursor
from PyQt6.QtCore import Qt, QSize, QPointF, QLineF, QRectF, pyqtSignal, QByteArray, QBuffer, QIODevice

class ThumbnailListWidget(QListWidget):
    """幅に合わせてアイコンサイズを自動調整するサムネイルリスト"""
    def __init__(self, parent=None):
        super().__init__(parent)
        # 横スクロールバーを出さないようにする
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # ビューポートの幅から余白を引いて新しい幅を計算
        w = max(50, self.viewport().width() - 25)
        h = int(w * 1.414) # A4用紙に近いアスペクト比
        
        # アイコンのサイズを更新
        self.setIconSize(QSize(w, h))
        # 文字の高さを含めたマスのサイズ(グリッド)を更新
        self.setGridSize(QSize(w + 5, h + 25))

class ResizeHandle(QGraphicsObject):
    """リサイズ用のハンドル（四隅と四辺の中点）"""
    def __init__(self, pos_type, parent):
        super().__init__(parent)
        self.pos_type = pos_type
        self.setAcceptHoverEvents(True)
        # ズームしてもハンドルの見た目のサイズを一定に保つ
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIgnoresTransformations)
        self.size = 8
        self._rect = QRectF(-self.size/2, -self.size/2, self.size, self.size)
        
        # 位置に応じたカーソル形状の設定
        if pos_type in ('top_left', 'bottom_right'):
            self.setCursor(Qt.CursorShape.SizeFDiagCursor)
        elif pos_type in ('top_right', 'bottom_left'):
            self.setCursor(Qt.CursorShape.SizeBDiagCursor)
        elif pos_type in ('top', 'bottom'):
            self.setCursor(Qt.CursorShape.SizeVerCursor)
        elif pos_type in ('left', 'right'):
            self.setCursor(Qt.CursorShape.SizeHorCursor)

    def boundingRect(self):
        return self._rect

    def paint(self, painter, option, widget):
        painter.setPen(QPen(QColor(0, 0, 0)))
        painter.setBrush(QColor(255, 255, 255))
        painter.drawRect(self._rect)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.parentItem().handle_mouse_press(self.pos_type, event.scenePos())
            event.accept()

    def mouseMoveEvent(self, event):
        self.parentItem().handle_mouse_move(self.pos_type, event.scenePos())
        event.accept()

    def mouseReleaseEvent(self, event):
        self.parentItem().handle_mouse_release(self.pos_type, event.scenePos())
        event.accept()

class ResizableRectItem(QGraphicsObject):
    """リサイズ・移動可能な選択範囲矩形"""
    rect_changed = pyqtSignal(QRectF)

    def __init__(self, rect=QRectF()):
        super().__init__()
        self._rect = rect
        self.setAcceptHoverEvents(True)
        self.setZValue(100)
        
        # 8つのハンドルを生成
        self.handles = {}
        for pos_type in ['top_left', 'top', 'top_right', 'right', 'bottom_right', 'bottom', 'bottom_left', 'left']:
            handle = ResizeHandle(pos_type, self)
            self.handles[pos_type] = handle
            
        self.update_handles()
        
        self._is_resizing = False
        self._resize_handle = None
        self._is_moving = False
        self._move_start_pos = QPointF()
        self._move_start_rect = QRectF()

    def boundingRect(self):
        # ハンドルが少しはみ出す分をマージンとして取る
        return self._rect.adjusted(-10, -10, 10, 10)

    def paint(self, painter, option, widget):
        pen = QPen(QColor(255, 0, 0), 2, Qt.PenStyle.DashLine)
        pen.setCosmetic(True) # ズームに関わらず線の太さを1px一定にする
        painter.setPen(pen)
        painter.setBrush(QColor(255, 0, 0, 40))
        painter.drawRect(self._rect)

    def rect(self):
        return self._rect
        
    def setRect(self, rect):
        self.prepareGeometryChange()
        self._rect = rect
        self.update_handles()
        self.update()

    def update_handles(self):
        r = self._rect
        self.handles['top_left'].setPos(r.topLeft())
        self.handles['top'].setPos(r.left() + r.width()/2, r.top())
        self.handles['top_right'].setPos(r.topRight())
        self.handles['right'].setPos(r.right(), r.top() + r.height()/2)
        self.handles['bottom_right'].setPos(r.bottomRight())
        self.handles['bottom'].setPos(r.left() + r.width()/2, r.bottom())
        self.handles['bottom_left'].setPos(r.bottomLeft())
        self.handles['left'].setPos(r.left(), r.top() + r.height()/2)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._is_moving = True
            self._move_start_pos = event.scenePos()
            self._move_start_rect = self._rect
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._is_moving:
            delta = event.scenePos() - self._move_start_pos
            new_rect = self._move_start_rect.translated(delta)
            self.setRect(new_rect)
            event.accept()
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self._is_moving:
            self._is_moving = False
            self.rect_changed.emit(self._rect)
            event.accept()
        else:
            super().mouseReleaseEvent(event)

    def hoverEnterEvent(self, event):
        self.setCursor(Qt.CursorShape.SizeAllCursor)
        super().hoverEnterEvent(event)
        
    def hoverLeaveEvent(self, event):
        self.setCursor(Qt.CursorShape.ArrowCursor)
        super().hoverLeaveEvent(event)

    def handle_mouse_press(self, handle_type, pos):
        self._is_resizing = True
        self._resize_handle = handle_type
        self._move_start_rect = self._rect

    def handle_mouse_move(self, handle_type, pos):
        if not self._is_resizing:
            return
        
        r = QRectF(self._move_start_rect)
        
        if 'left' in handle_type:
            r.setLeft(min(pos.x(), r.right() - 10))
        elif 'right' in handle_type:
            r.setRight(max(pos.x(), r.left() + 10))
            
        if 'top' in handle_type:
            r.setTop(min(pos.y(), r.bottom() - 10))
        elif 'bottom' in handle_type:
            r.setBottom(max(pos.y(), r.top() + 10))
            
        self.setRect(r.normalized())

    def handle_mouse_release(self, handle_type, pos):
        self._is_resizing = False
        self._resize_handle = None
        self.rect_changed.emit(self._rect)

class CustomGraphicsView(QGraphicsView):
    """プレビューキャンバス用のカスタムビュー（ズーム、パン、矩形選択機能付き）"""
    
    # 矩形選択が完了したときに発行するシグナル
    crop_rect_completed = pyqtSignal(QRectF)
    
    def __init__(self, scene):
        super().__init__(scene)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.DragMode.NoDrag)
        
        # 背景をグレーに変更
        self.setBackgroundBrush(QBrush(QColor(85, 85, 85))) # #555555
        
        self.show_grid = True
        self.current_tool = "rect" # "pen", "eraser", "rect"
        
        self._is_panning = False
        self._pan_start_pos = QPointF()
        
        self._is_drawing_rect = False
        self._rect_start = QPointF()
        self._drawing_rect_item = None

    def drawForeground(self, painter, rect):
        """キャンバス前面にグリッドを描画する"""
        super().drawForeground(painter, rect)
        if not getattr(self, 'show_grid', False):
            return

        painter.save()
        painter.setTransform(QTransform())
        pen = QPen(QColor(200, 200, 200, 150), 1, Qt.PenStyle.DashLine)
        painter.setPen(pen)
        
        viewport_rect = self.viewport().rect()
        width = viewport_rect.width()
        height = viewport_rect.height()
        grid_size = 50
        
        lines = []
        for x in range(0, width, grid_size):
            lines.append(QLineF(x, 0, x, height))
        for y in range(0, height, grid_size):
            lines.append(QLineF(0, y, width, y))
            
        painter.drawLines(lines)
        painter.restore()

    def wheelEvent(self, event):
        """Ctrl + マウスホイールでズームイン/ズームアウト"""
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            zoom_in_factor = 1.15
            zoom_out_factor = 1.0 / zoom_in_factor
            old_pos = self.mapToScene(event.position().toPoint())
            
            if event.angleDelta().y() > 0:
                zoom_factor = zoom_in_factor
            else:
                zoom_factor = zoom_out_factor
                
            self.scale(zoom_factor, zoom_factor)
            new_pos = self.mapToScene(event.position().toPoint())
            delta = new_pos - old_pos
            self.translate(delta.x(), delta.y())
        else:
            super().wheelEvent(event)

    def mousePressEvent(self, event):
        """マウスクリック処理（パンと矩形選択）"""
        if event.button() == Qt.MouseButton.MiddleButton:
            self._is_panning = True
            self._pan_start_pos = event.position()
            self.setCursor(Qt.CursorShape.ClosedHandCursor)
            event.accept()
            
        elif event.button() == Qt.MouseButton.LeftButton and self.current_tool == "rect":
            # 既存のハンドルや矩形の上をクリックした場合は、移動・リサイズイベントを優先
            item = self.itemAt(event.position().toPoint())
            if isinstance(item, (ResizableRectItem, ResizeHandle)):
                super().mousePressEvent(event)
                return

            self._is_drawing_rect = True
            self._rect_start = self.mapToScene(event.position().toPoint())
            
            if self._drawing_rect_item:
                self.scene().removeItem(self._drawing_rect_item)
                
            self._drawing_rect_item = QGraphicsRectItem()
            pen = QPen(QColor(255, 0, 0), 2, Qt.PenStyle.DashLine)
            self._drawing_rect_item.setPen(pen)
            self._drawing_rect_item.setBrush(QColor(255, 0, 0, 40)) # 半透明の赤
            self._drawing_rect_item.setZValue(100) # 画像より手前に表示
            self.scene().addItem(self._drawing_rect_item)
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        """マウス移動処理"""
        if self._is_panning:
            delta = event.position() - self._pan_start_pos
            self.horizontalScrollBar().setValue(int(self.horizontalScrollBar().value() - delta.x()))
            self.verticalScrollBar().setValue(int(self.verticalScrollBar().value() - delta.y()))
            self._pan_start_pos = event.position()
            event.accept()
            
        elif self._is_drawing_rect:
            current_pos = self.mapToScene(event.position().toPoint())
            rect = QRectF(self._rect_start, current_pos).normalized()
            self._drawing_rect_item.setRect(rect)
            event.accept()
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        """マウスリリース処理"""
        if event.button() == Qt.MouseButton.MiddleButton:
            self._is_panning = False
            self.setCursor(QCursor(Qt.CursorShape.CrossCursor) if self.current_tool == "rect" else Qt.CursorShape.ArrowCursor)
            event.accept()
            
        elif event.button() == Qt.MouseButton.LeftButton and self._is_drawing_rect:
            self._is_drawing_rect = False
            if self._drawing_rect_item:
                final_rect = self._drawing_rect_item.rect()
                # 仮描画用アイテムは一度削除する（シグナル先でResizableRectItemとして再作成される）
                self.scene().removeItem(self._drawing_rect_item)
                self._drawing_rect_item = None
                
                if final_rect.width() > 10 and final_rect.height() > 10:
                    self.crop_rect_completed.emit(final_rect)
            event.accept()
        else:
            super().mouseReleaseEvent(event)


class PDFAdjuster(QMainWindow):
    """メインアプリケーションウィンドウ"""
    
    # 出力用紙のサイズ定義 (幅, 高さ) px設定 (約300dpi)
    PAPER_SIZES = {
        "A4": (2480, 3508),
        "B5": (2079, 2953),
        "A5": (1748, 2480),
        "Letter": (2550, 3300)
    }

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Adjuster - 範囲選択と正規化")
        self.setGeometry(100, 100, 1200, 800)
        
        self.pdf_doc = None
        # 各ページの状態: angle(傾き), crop_rect(選択範囲 QRectF)
        self.page_states = {} 
        self.current_page_num = -1
        self.image_item = None
        self.crop_rect_item = None

        self._create_menu()
        self._init_ui()

    def _create_menu(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu("ファイル(&F)")
        
        open_action = QAction("PDFを開く(&O)", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_pdf_dialog)
        file_menu.addAction(open_action)

    def _init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(main_splitter)
        
        # === 左パネル: サムネイル一覧 ===
        self.thumbnail_list = ThumbnailListWidget()
        self.thumbnail_list.setMinimumWidth(120)
        self.thumbnail_list.setViewMode(QListWidget.ViewMode.IconMode)
        self.thumbnail_list.setMovement(QListWidget.Movement.Static)
        self.thumbnail_list.setResizeMode(QListWidget.ResizeMode.Adjust)
        
        # 常に縦1列で表示し、横に折り返さない設定
        self.thumbnail_list.setFlow(QListView.Flow.TopToBottom)
        self.thumbnail_list.setWrapping(False)
        
        self.thumbnail_list.itemSelectionChanged.connect(self.on_page_selected)
        main_splitter.addWidget(self.thumbnail_list)
        
        # === 中央パネル: プレビューキャンバス ===
        self.scene = QGraphicsScene()
        self.preview_view = CustomGraphicsView(self.scene)
        self.preview_view.crop_rect_completed.connect(self.on_crop_rect_drawn)
        self.preview_view.setCursor(Qt.CursorShape.CrossCursor) # 初期カーソル
        main_splitter.addWidget(self.preview_view)
        
        # === 右パネル: コントロール ===
        right_panel = QWidget()
        right_panel.setMinimumWidth(250)
        control_layout = QVBoxLayout(right_panel)
        
        # PDFを開くボタン
        self.btn_open = QPushButton("PDFを開く")
        self.btn_open.clicked.connect(self.open_pdf_dialog)
        self.btn_open.setStyleSheet("padding: 8px; font-weight: bold;")
        control_layout.addWidget(self.btn_open)
        control_layout.addSpacing(10)
        
        # 傾きコントロール
        control_layout.addWidget(QLabel("傾き調整 (度):"))
        angle_layout = QHBoxLayout()
        self.angle_spinbox = QDoubleSpinBox()
        self.angle_spinbox.setRange(-180.0, 180.0)
        self.angle_spinbox.setSingleStep(0.1)
        self.angle_spinbox.setDecimals(1)
        self.angle_spinbox.valueChanged.connect(self._on_spinbox_changed)
        angle_layout.addWidget(self.angle_spinbox)
        
        self.angle_slider = QSlider(Qt.Orientation.Horizontal)
        self.angle_slider.setRange(-1800, 1800)
        self.angle_slider.valueChanged.connect(self._on_slider_changed)
        angle_layout.addWidget(self.angle_slider)
        control_layout.addLayout(angle_layout)
        
        # グリッド表示トグル
        self.grid_checkbox = QCheckBox("グリッドを表示")
        self.grid_checkbox.setChecked(True)
        self.grid_checkbox.stateChanged.connect(self._on_grid_toggled)
        control_layout.addWidget(self.grid_checkbox)
        
        # ツール類
        control_layout.addSpacing(20)
        control_layout.addWidget(QLabel("ツール:"))
        self.pen_btn = QPushButton("ペン")
        self.eraser_btn = QPushButton("消しゴム")
        self.select_btn = QPushButton("コンテンツ範囲選択")
        self.select_btn.setChecked(True)
        
        self.tool_group = QButtonGroup()
        self.tool_group.addButton(self.pen_btn, 1)
        self.tool_group.addButton(self.eraser_btn, 2)
        self.tool_group.addButton(self.select_btn, 3)
        self.pen_btn.setCheckable(True)
        self.eraser_btn.setCheckable(True)
        self.select_btn.setCheckable(True)
        
        # ツール切り替えイベント
        self.tool_group.idToggled.connect(self.on_tool_changed)
        
        control_layout.addWidget(self.pen_btn)
        control_layout.addWidget(self.eraser_btn)
        control_layout.addWidget(self.select_btn)
        
        # 出力設定
        control_layout.addSpacing(40)
        control_layout.addWidget(QLabel("出力用紙サイズ:"))
        self.paper_size_combo = QComboBox()
        self.paper_size_combo.addItems(list(self.PAPER_SIZES.keys()))
        control_layout.addWidget(self.paper_size_combo)

        # エクスポートボタン
        self.btn_export = QPushButton("PDFエクスポート実行")
        self.btn_export.setStyleSheet("background-color: #4CAF50; color: white; padding: 12px; font-weight: bold;")
        self.btn_export.clicked.connect(self.export_pdf)
        control_layout.addWidget(self.btn_export)
        
        control_layout.addStretch()
        main_splitter.addWidget(right_panel)
        
        # スプリッターの初期サイズ比率を設定（左パネル, 中央キャンバス, 右パネル）
        main_splitter.setSizes([180, 800, 250])

    def on_tool_changed(self, id, checked):
        if checked:
            if id == 3:
                self.preview_view.current_tool = "rect"
                self.preview_view.setCursor(Qt.CursorShape.CrossCursor)
            elif id == 1:
                self.preview_view.current_tool = "pen"
                self.preview_view.setCursor(Qt.CursorShape.ArrowCursor)
            elif id == 2:
                self.preview_view.current_tool = "eraser"
                self.preview_view.setCursor(Qt.CursorShape.ArrowCursor)

    def on_crop_rect_drawn(self, rect):
        """キャンバスで新規矩形が描画完了したときの処理"""
        if self.current_page_num != -1:
            self.page_states[self.current_page_num]['crop_rect'] = rect
            
            if self.crop_rect_item:
                self.scene.removeItem(self.crop_rect_item)
                
            self.crop_rect_item = ResizableRectItem(rect)
            self.crop_rect_item.rect_changed.connect(self.on_crop_rect_modified)
            self.scene.addItem(self.crop_rect_item)

    def on_crop_rect_modified(self, new_rect):
        """リサイズハンドルや移動によって矩形が変更されたときの処理"""
        if self.current_page_num != -1:
            self.page_states[self.current_page_num]['crop_rect'] = new_rect

    def open_pdf_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "PDFを開く", "", "PDF Files (*.pdf)")
        if file_path:
            self.load_pdf(file_path)

    def load_pdf(self, file_path):
        if self.pdf_doc:
            self.pdf_doc.close()
            
        self.pdf_doc = fitz.open(file_path)
        self.page_states = {i: {'angle': 0.0, 'crop_rect': None} for i in range(len(self.pdf_doc))}
        self.current_page_num = -1
        
        self.thumbnail_list.clear()
        self.scene.clear()
        
        for i in range(len(self.pdf_doc)):
            page = self.pdf_doc[i]
            # アイコンが上限サイズで頭打ちにならないよう、生成解像度を大きくする(1.5倍)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
            qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
            qpix = QPixmap.fromImage(qimg)
            
            item = QListWidgetItem(QIcon(qpix), str(i+1))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setData(Qt.ItemDataRole.UserRole, i)
            self.thumbnail_list.addItem(item)
            
        if self.thumbnail_list.count() > 0:
            self.thumbnail_list.setCurrentRow(0)

    def on_page_selected(self):
        selected_items = self.thumbnail_list.selectedItems()
        if not selected_items:
            return
            
        item = selected_items[0]
        page_num = item.data(Qt.ItemDataRole.UserRole)
        
        if page_num == self.current_page_num:
            return
            
        self.current_page_num = page_num
        self.render_preview(page_num)
        
        state = self.page_states[page_num]
        
        self.angle_spinbox.blockSignals(True)
        self.angle_slider.blockSignals(True)
        self.angle_spinbox.setValue(state['angle'])
        self.angle_slider.setValue(int(state['angle'] * 10))
        self.angle_spinbox.blockSignals(False)
        self.angle_slider.blockSignals(False)

    def render_preview(self, page_num):
        self.scene.clear()
        self.preview_view._drawing_rect_item = None
        self.crop_rect_item = None
        
        page = self.pdf_doc[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(4.0, 4.0))
        fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
        qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
        qpix = QPixmap.fromImage(qimg)
        
        self.image_item = QGraphicsPixmapItem(qpix)
        self.image_item.setTransformOriginPoint(qpix.width() / 2, qpix.height() / 2)
        
        self.scene.addItem(self.image_item)
        self.scene.setSceneRect(0, 0, qpix.width(), qpix.height())
        
        # 記憶されている傾きを適用
        self.apply_rotation(self.page_states[page_num]['angle'])
        
        # 記憶されている選択枠を描画
        saved_rect = self.page_states[page_num].get('crop_rect')
        if saved_rect:
            self.crop_rect_item = ResizableRectItem(saved_rect)
            self.crop_rect_item.rect_changed.connect(self.on_crop_rect_modified)
            self.scene.addItem(self.crop_rect_item)
        
        self.preview_view.fitInView(self.scene.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)

    def _on_slider_changed(self, value):
        angle = value / 10.0
        self.angle_spinbox.blockSignals(True)
        self.angle_spinbox.setValue(angle)
        self.angle_spinbox.blockSignals(False)
        self._update_current_page_angle(angle)

    def _on_spinbox_changed(self, value):
        slider_val = int(value * 10)
        self.angle_slider.blockSignals(True)
        self.angle_slider.setValue(slider_val)
        self.angle_slider.blockSignals(False)
        self._update_current_page_angle(value)

    def _update_current_page_angle(self, angle):
        if self.current_page_num != -1:
            self.page_states[self.current_page_num]['angle'] = angle
            self.apply_rotation(angle)

    def _on_grid_toggled(self, state):
        self.preview_view.show_grid = (state != 0)
        self.preview_view.viewport().update()

    def apply_rotation(self, angle):
        if self.image_item:
            self.image_item.setRotation(angle)

    def export_pdf(self):
        """設定を適用し、位置合わせ・余白白塗りを行った新しいPDFを出力する"""
        if not self.pdf_doc:
            QMessageBox.warning(self, "エラー", "PDFが読み込まれていません。")
            return
            
        save_path, _ = QFileDialog.getSaveFileName(self, "PDFを保存", "adjusted_document.pdf", "PDF Files (*.pdf)")
        if not save_path:
            return
            
        # 選択された用紙サイズを取得
        paper_name = self.paper_size_combo.currentText()
        target_w, target_h = self.PAPER_SIZES[paper_name]
        
        out_pdf = fitz.open()
        
        # プログレス表示の代わり
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        
        try:
            for i in range(len(self.pdf_doc)):
                state = self.page_states[i]
                angle = state['angle']
                crop_rect = state.get('crop_rect')
                
                # 1. 元ページを高解像度で取得 (プレビューと同じ基準)
                page = self.pdf_doc[i]
                pix = page.get_pixmap(matrix=fitz.Matrix(4.0, 4.0))
                fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
                qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
                
                # 2. 出力用の白紙画像を作成（外側を白くするベース）
                out_img = QImage(target_w, target_h, QImage.Format.Format_RGB32)
                out_img.fill(Qt.GlobalColor.white)
                painter = QPainter(out_img)
                painter.setRenderHint(QPainter.RenderHint.Antialiasing)
                painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
                
                # 3. スケールとオフセットの計算
                if crop_rect and not crop_rect.isEmpty():
                    # 仕様: 選択範囲の縦の長さが用紙の80%になるように揃える
                    required_h = target_h * 0.8
                    scale = required_h / crop_rect.height()
                    cx = crop_rect.center().x()
                    cy = crop_rect.center().y()
                else:
                    # 範囲指定がない場合は適当に全体を収める
                    scale = min(target_w / qimg.width(), target_h / qimg.height()) * 0.9
                    cx = qimg.width() / 2
                    cy = qimg.height() / 2

                # 4. 白紙画像の中心に原点を移動
                painter.translate(target_w / 2, target_h / 2)
                painter.scale(scale, scale)
                
                # 5. 選択範囲外を白にするためのクリッピング
                if crop_rect and not crop_rect.isEmpty():
                    # 現在の原点がcx, cyと対応しているので、中心基準の矩形でクリップ
                    clip_w, clip_h = crop_rect.width(), crop_rect.height()
                    painter.setClipRect(QRectF(-clip_w/2, -clip_h/2, clip_w, clip_h))
                
                # 6. 元画像を描画位置へオフセット
                painter.translate(-cx, -cy)
                
                # 7. 元画像自体の回転を適用
                img_cx, img_cy = qimg.width() / 2, qimg.height() / 2
                painter.translate(img_cx, img_cy)
                painter.rotate(angle)
                painter.translate(-img_cx, -img_cy)
                
                painter.drawImage(0, 0, qimg)
                painter.end()
                
                # QImage を jpegバイト列経由で PDFページ に変換して追加
                ba = QByteArray()
                buffer = QBuffer(ba)
                buffer.open(QIODevice.OpenModeFlag.WriteOnly)
                out_img.save(buffer, "JPEG", quality=95)
                
                img_pdf = fitz.open("pdf", fitz.open("jpeg", ba.data()).convert_to_pdf())
                out_pdf.insert_pdf(img_pdf)
                
            out_pdf.save(save_path)
            out_pdf.close()
            QMessageBox.information(self, "完了", "PDFのエクスポートが完了しました。")
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"出力中にエラーが発生しました:\n{e}")
        finally:
            QApplication.restoreOverrideCursor()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PDFAdjuster()
    window.show()
    sys.exit(app.exec())