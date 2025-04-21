# -*- coding: utf-8 -*-
import sys
import os
import pandas as pd
import numpy as np
import re
from difflib import SequenceMatcher
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit,
                            QWidget, QMessageBox)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Автоматизация спецификаций")
        self.setGeometry(100, 100, 800, 600)
        
        # Основные переменные
        self.sp_df = None
        self.ndb_df = None
        self.output_dir = ""
        self.product_name = ""
        
        # Создаем главный виджет
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        
        # Создаем макет
        self.layout = QVBoxLayout()
        self.main_widget.setLayout(self.layout)
        
        # Добавляем элементы интерфейса
        self.create_ui()
        
        # Инициализация словарей
        self.init_dictionaries()
        
    def create_ui(self):
        # Поле для имени изделия
        self.layout.addWidget(QLabel("Введите наименование изделия:"))
        self.product_name_input = QLineEdit()
        self.layout.addWidget(self.product_name_input)
        
        # Кнопка загрузки спецификации
        self.spec_button = QPushButton("Выберите выгрузку спецификации из SolidWorks")
        self.spec_button.clicked.connect(self.load_specification)
        self.layout.addWidget(self.spec_button)
        
        # Кнопка загрузки номенклатуры
        self.nomenclature_button = QPushButton("Выберите файл с номенклатурой")
        self.nomenclature_button.clicked.connect(self.load_nomenclature)
        self.layout.addWidget(self.nomenclature_button)
        
        # Консоль вывода
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.layout.addWidget(self.console)
        
        # Кнопка обработки
        self.process_button = QPushButton("ОК")
        self.process_button.clicked.connect(self.process_data)
        self.layout.addWidget(self.process_button)
        
        # Кнопки сохранения
        button_layout = QHBoxLayout()
        
        self.save_purchase_button = QPushButton("Сохранить спецификацию на закупку")
        self.save_purchase_button.clicked.connect(self.save_purchase_spec)
        self.save_purchase_button.setEnabled(False)
        button_layout.addWidget(self.save_purchase_button)
        
        self.save_tech_button = QPushButton("Сохранить техпроцессы")
        self.save_tech_button.clicked.connect(self.save_tech_processes)
        self.save_tech_button.setEnabled(False)
        button_layout.addWidget(self.save_tech_button)
        
        self.layout.addLayout(button_layout)
    
    def init_dictionaries(self):
        # Словарь крепежа
        self.sp_dict = {
            "Крепеж": [
                "Болт", "Винт", "Шпилька", "Гайка", "Шайба", 
                "Резьбовая заклепка", "Евровинт", "Заклепка", 
                "Штифт", "Шплинт", "Стопорное кольцо", "Клин", 
                "Гвоздь", "Скоба", "Анкер", "Дюбель", "Хомут"
            ]
        }
        
        # Словарь расшифровок кодов
        self.code_decoder = {
            'АП': '3D печать', 'ЛР': 'Лазерная резка', 'Г': 'Гибка',
            'ВЛ': 'Вальцовка', 'МО': 'Мех обработка', 'ФР': 'Фрезеровка',
            'ТО': 'Токарная обработка', 'ШЛ': 'Шлифовка',
            'СД': 'Сварка деталей (в процессе сборки)',
            'СШ': 'Сварка шва (внутри детали)', 'ВР': 'Воронение',
            'ПК': 'Покраска', 'ЗК': 'Закалка', 'ТОБ': 'Термообработка',
            'АО': 'Абразивная обработка', 'ХО': 'Химическая обработка',
            'СЛ': 'Слесарные работы', 'НП': 'Нанесение покрытий',
            'ГА': 'Гидроабразив'
        }
        
        # Типы операций
        self.tech_processes = ['АП', 'ЛР', 'Г', 'ВЛ', 'МО', 'ФР', 'ТО', 'ГА']
        self.welding_ops = ['СД', 'СШ']
        self.post_processing = ['ШЛ', 'ВР', 'ПК', 'ЗК', 'ТОБ', 'АО', 'ХО', 'СЛ', 'НП']
    
    def log_message(self, message):
        """Вывод сообщения в консоль"""
        self.console.append(message)
        QApplication.processEvents()  # Обновляем интерфейс
    
    def load_specification(self):
        """Загрузка файла спецификации"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл спецификации", 
            "", "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)")
        
        if file_path:
            self.output_dir = os.path.dirname(file_path)
            try:
                if file_path.endswith('.csv'):
                    self.sp_df = pd.read_csv(file_path)
                else:
                    self.sp_df = pd.read_excel(file_path)
                
                self.log_message(f"Загружен файл спецификации: {os.path.basename(file_path)}")
                self.log_message(f"Загружено строк: {len(self.sp_df)}")
                
            except Exception as e:
                self.log_message(f"Ошибка при загрузке файла: {str(e)}")
                QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл: {str(e)}")
    
    def load_nomenclature(self):
        """Загрузка файла номенклатуры"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл номенклатуры", 
            "", "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)")
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.ndb_df = pd.read_csv(file_path)
                else:
                    self.ndb_df = pd.read_excel(file_path)
                
                self.log_message(f"Загружен файл номенклатуры: {os.path.basename(file_path)}")
                self.log_message(f"Загружено строк: {len(self.ndb_df)}")
                
            except Exception as e:
                self.log_message(f"Ошибка при загрузке файла: {str(e)}")
                QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл: {str(e)}")
    
    def process_data(self):
        """Основная обработка данных"""
        if self.sp_df is None:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите файл спецификации!")
            return
            
        self.product_name = self.product_name_input.text().strip()
        if not self.product_name:
            QMessageBox.warning(self, "Ошибка", "Введите наименование изделия!")
            return
            
        try:
            self.log_message("\nНачало обработки данных...")
            
            # Создаем копию DataFrame
            sp_df_f = self.sp_df.copy()
            
            # 1. Очистка данных
            self.log_message("Очистка данных...")
            self.clean_data(sp_df_f)
            
            # 2. Добавление столбцов
            self.log_message("Добавление столбцов...")
            self.add_columns(sp_df_f)
            
            # 3. Формирование примечания
            self.log_message("Формирование примечаний...")
            self.process_notes(sp_df_f)
            
            # 4. Работа с крепежом
            self.log_message("Обработка крепежа...")
            self.process_fasteners(sp_df_f)
            
            # 5. Агрегация дубликатов
            self.log_message("Агрегация дубликатов...")
            sp_df_f = self.aggregate_duplicates(sp_df_f)
            
            # 6. Замена групп закупки
            self.log_message("Обновление групп закупки...")
            self.update_purchase_groups(sp_df_f)
            
            # 7. Анализ номенклатуры
            if self.ndb_df is not None:
                self.log_message("Анализ номенклатуры...")
                self.analyze_nomenclature(sp_df_f)
            
            # Фильтрация итоговых данных
            self.log_message("Фильтрация данных...")
            self.sp_df_f_clear = sp_df_f[~sp_df_f["Группа закупки"].isin(
                [0, "Модель", "Сборка по месту", np.nan, ""])]
            
            # Создание техкарты
            self.log_message("Формирование техкарты...")
            self.create_tech_card(sp_df_f)
            
            self.log_message("\nОбработка завершена успешно!")
            
            # Активируем кнопки сохранения
            self.save_purchase_button.setEnabled(True)
            self.save_tech_button.setEnabled(True)
            
        except Exception as e:
            self.log_message(f"\nОшибка при обработке данных: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при обработке данных: {str(e)}")
    
    def clean_data(self, df):
        """Очистка данных от лишних символов"""
        for column in df.columns:
            if df[column].dtype == 'object':
                try:
                    df[column] = df[column].str.replace(r'\n', '', regex=True)
                    df[column] = df[column].str.strip()
                except AttributeError:
                    pass
    
    def add_columns(self, df):
        """Добавление необходимых столбцов"""
        columns_to_add = [
            (4, "Цена руб. с НДС"),
            (5, "Сумма"),
            (6, "Срок, нед."),
            (10, "Поставщик")
        ]
        
        for idx, name in columns_to_add:
            if name not in df.columns:
                df.insert(idx, name, "")
        
        df.fillna("", inplace=True)
    
    def process_notes(self, df):
        """Формирование примечания материалов деталей"""
        target_groups = ["Лазерная резка", "Мех. обработка", "Фрезеровка", 
                        "Токарная обработка", "Гидроабразив", "3D печать"]
        
        mask = df["Группа закупки"].isin(target_groups)
        
        def combine_material_zagotovka(row):
            material = str(row["Материал"])
            zagotovka = str(row["Заготовка"])
            return f"{material} - {zagotovka}" if zagotovka.strip() else material
        
        df.loc[mask, "Примечание"] = df.loc[mask].apply(combine_material_zagotovka, axis=1)
        
        if "Материал" in df.columns:
            df.drop("Материал", axis=1, inplace=True)
        if "Заготовка" in df.columns:
            df.drop("Заготовка", axis=1, inplace=True)
    
    def process_fasteners(self, df):
        """Обработка крепежа"""
        # Замена GB_SPECIAL_TYPE
        def replace_gb_special(row):
            if 'GB_SPECIAL_TYPE' in row['Наименование']:
                parts = row['Наименование'].split(' ')
                if len(parts) > 1:
                    after_gb = parts[1]
                    x_index = after_gb.upper().find('X')
                    return 'Заклепка резьбовая ' + (after_gb[:x_index] if x_index != -1 else after_gb)
                return 'Заклепка резьбовая'
            return row['Наименование']
        
        df['Наименование'] = df.apply(replace_gb_special, axis=1)
        
        # Обновление группы закупки для крепежа
        def update_group(row):
            for item in self.sp_dict['Крепеж']:
                if item in row['Наименование']:
                    return 'Крепеж'
            return row['Группа закупки']
        
        df['Группа закупки'] = df.apply(update_group, axis=1)
    
    def aggregate_duplicates(self, df):
        """Агрегация дубликатов"""
        original_columns = df.columns.tolist()
        
        df["Кол-во"] = pd.to_numeric(df["Кол-во"], errors='coerce').fillna(0)
        df["Примечание"] = df["Примечание"].fillna("")
        
        grouped = df.groupby("Наименование")
        
        aggregated_df = grouped.agg({
            "ПОЗИЦИЯ": "first",
            "Артикул": "first",
            "Кол-во": "sum",
            "Цена руб. с НДС": "first",
            "Сумма": "first",
            "Срок, нед.": "first",
            "Примечание": lambda x: max(x, key=len),
            "Поставщик": "first",
            "Производитель": "first",
            "Группа закупки": "first",
        }).reset_index()
        
        return aggregated_df[original_columns]
    
    def update_purchase_groups(self, df):
        """Обновление групп закупки"""
        def first_tech(article):
            if pd.isna(article):
                return "Деталь"
                
            parts = str(article).split('.')
            for part in parts:
                if part in self.tech_processes:
                    return self.code_decoder[part]
            return "Деталь"
        
        mask = df['Группа закупки'] == 'Деталь'
        df.loc[mask, 'Группа закупки'] = df.loc[mask, 'Артикул'].apply(first_tech)
    
    def analyze_nomenclature(self, df):
        """Анализ номенклатуры"""
        required_columns = ['Наименование', 'Артикул', 'Производитель', 'Комментарий']
        if not all(col in self.ndb_df.columns for col in required_columns):
            self.log_message("В базе номенклатуры отсутствуют необходимые столбцы")
            return
            
        self.ndb_df = self.ndb_df.fillna("")
        
        def normalize_article(article):
            return re.sub(r'[_\-\s/]', '', str(article).lower())
        
        def find_best_match(name, ndb_df):
            normalized_name = normalize_article(name)
            best_match = None
            best_score = 0
            
            for _, row in ndb_df.iterrows():
                ndb_article = str(row['Артикул']).strip()
                if not ndb_article:
                    continue
                    
                normalized_ndb = normalize_article(ndb_article)
                
                if normalized_ndb in normalized_name:
                    return row
                
                current_score = SequenceMatcher(None, normalized_ndb, normalized_name).ratio()
                if current_score > best_score and current_score >= 0.9:
                    best_score = current_score
                    best_match = row
            
            return best_match if best_score >= 0.9 else None
        
        updated_count = 0
        
        for idx, row in df.iterrows():
            name = str(row['Наименование'])
            match = find_best_match(name, self.ndb_df)
            
            if match is not None:
                df.at[idx, 'Наименование'] = match['Наименование']
                df.at[idx, 'Артикул'] = match['Артикул']
                df.at[idx, 'Производитель'] = match['Производитель']
                df.at[idx, 'Группа закупки'] = match['Комментарий']
                updated_count += 1
        
        self.log_message(f"Обновлено позиций из номенклатуры: {updated_count}/{len(df)}")
    
    def create_tech_card(self, df):
        """Создание техкарты"""
        target_groups = ['Деталь', 'Лазерная резка', 'Мех обработка', 
                        'Фрезеровка', 'Токарная обработка', 'Сварочная сборка']
        
        self.df_teh = df[df['Группа закупки'].isin(target_groups)][
            ['Наименование', 'Артикул', 'Группа закупки']].copy()
        
        # Обработка техпроцессов
        def parse_article(article):
            if pd.isna(article):
                return [], [], []
                
            parts = [part.lower() for part in str(article).split('.')]
            techs, welds, posts = [], [], []
            
            for part in parts:
                if part in [t.lower() for t in self.tech_processes]:
                    techs.append(self.code_decoder[part.upper()])
                elif part in [w.lower() for w in self.welding_ops]:
                    welds.append(self.code_decoder[part.upper()])
                elif part in [p.lower() for p in self.post_processing]:
                    posts.append(self.code_decoder[part.upper()])
            
            return techs, welds, posts
        
        # Анализ артикулов
        max_tech = max_weld = max_post = 0
        tech_data = []
        
        for article in self.df_teh['Артикул']:
            techs, welds, posts = parse_article(article)
            tech_data.append((techs, welds, posts))
            max_tech = max(max_tech, len(techs))
            max_weld = max(max_weld, len(welds))
            max_post = max(max_post, len(posts))
        
        # Добавление столбцов
        for i in range(1, max_tech + 1):
            self.df_teh[f'Техпроцесс {i}'] = ''
        for i in range(1, max_weld + 1):
            self.df_teh[f'Сварочная операция {i}'] = ''
        for i in range(1, max_post + 1):
            self.df_teh[f'Постобработка {i}'] = ''
        
        # Заполнение данных
        for idx, (techs, welds, posts) in enumerate(tech_data):
            for i, tech in enumerate(techs, 1):
                self.df_teh.at[idx, f'Техпроцесс {i}'] = tech
            for i, weld in enumerate(welds, 1):
                self.df_teh.at[idx, f'Сварочная операция {i}'] = weld
            for i, post in enumerate(posts, 1):
                self.df_teh.at[idx, f'Постобработка {i}'] = post
        
        # Удаление лишних столбцов
        self.df_teh.drop(['Артикул', 'Группа закупки'], axis=1, inplace=True, errors='ignore')
    
    def save_purchase_spec(self):
        """Сохранение спецификации на закупку"""
        if not hasattr(self, 'sp_df_f_clear'):
            QMessageBox.warning(self, "Ошибка", "Сначала обработайте данные!")
            return
            
        if not self.output_dir:
            QMessageBox.warning(self, "Ошибка", "Не определена папка для сохранения!")
            return
            
        filename = f"Спецификация на закупку {self.product_name}.xlsx"
        file_path = os.path.join(self.output_dir, filename)
        
        try:
            self.save_excel_auto_width(self.sp_df_f_clear, file_path)
            self.log_message(f"\nФайл сохранен: {file_path}")
            QMessageBox.information(self, "Успех", f"Файл успешно сохранен:\n{file_path}")
        except Exception as e:
            self.log_message(f"\nОшибка при сохранении: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {str(e)}")
    
    def save_tech_processes(self):
        """Сохранение техпроцессов"""
        if not hasattr(self, 'df_teh'):
            QMessageBox.warning(self, "Ошибка", "Сначала обработайте данные!")
            return
            
        if not self.output_dir:
            QMessageBox.warning(self, "Ошибка", "Не определена папка для сохранения!")
            return
            
        filename = f"Техпроцессы {self.product_name}.xlsx"
        file_path = os.path.join(self.output_dir, filename)
        
        try:
            self.save_excel_auto_width(self.df_teh, file_path)
            self.log_message(f"\nФайл сохранен: {file_path}")
            QMessageBox.information(self, "Успех", f"Файл успешно сохранен:\n{file_path}")
        except Exception as e:
            self.log_message(f"\nОшибка при сохранении: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {str(e)}")
    
    def save_excel_auto_width(self, df, filename):
        """Сохранение Excel с авто-шириной столбцов"""
        wb = Workbook()
        ws = wb.active

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top')

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter

            for cell in col:
                try:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass

            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width

        wb.save(filename)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())