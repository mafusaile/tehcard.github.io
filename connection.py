import sqlite3 as sql
from tkinter.messagebox import showinfo, showerror
import os

if os.path.exists('КАЛЬКУЛЯЦИОННЫЕ КАРТЫ') == False:
	os.mkdir('КАЛЬКУЛЯЦИОННЫЕ КАРТЫ')

class Data:
	db = sql.connect('data.db')
	cursor = db.cursor()

	def __init__(self):
		super(Data, self).__init__()

	def create_database(self):
			create_table_ingredients = ("""CREATE TABLE IF NOT EXISTS ingredients(id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
					ingredient TEXT,
					weight NUMERIC,
					cost NUMERIC,
					protein NUMERIC,
					fats NUMERIC,
					carbohydrates NUMERIC,
					calorie_content NUMERIC,
					remainder NUMERIC
					);
			""")
			create_table_products_description = ("""
							CREATE TABLE IF NOT EXISTS products_description(id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
								product_name TEXT,
								appearance TEXT,
								consistency NUMERIC,
								taste NUMERIC,
								smell NUMERIC
								);
						""")
			create_table_products_composition = ("""
				CREATE TABLE IF NOT EXISTS products_composition(id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
					product_name TEXT,
					ingredient TEXT,
					quantity NUMERIC
					);
			""")
			Data.cursor.execute(create_table_ingredients)
			Data.cursor.execute(create_table_products_description)
			Data.cursor.execute(create_table_products_composition)

data = Data()
data.create_database()