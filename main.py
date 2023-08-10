import win32com.client
from datetime import datetime
import os
import zipfile
import tempfile
import shutil

ol = win32com.client.Dispatch("outlook.application")
now = datetime.now()
day = now.day
month = now.month

# Klasör adını oluştur
folder_name = f"{day:02d}.{month:02d}"
folder_path = os.path.join(r"C:\Users\...\Desktop", folder_name)

# Geçici klasör oluştur
temp_folder = tempfile.mkdtemp()

# Klasör içeriğini geçici klasöre kopyalar
for root, _, files in os.walk(folder_path):
    for file in files:
        file_path = os.path.join(root, file)
        shutil.copy2(file_path, os.path.join(temp_folder, file))  # Farklı bir geçici klasör kullan

# Arşiv dosyasının adını belirler
zip_file_name = "archive.zip"
zip_file_path = os.path.join(folder_path, zip_file_name)

# Temp klasörü zip dosyasına arşivler
with zipfile.ZipFile(zip_file_path, "w", zipfile.ZIP_DEFLATED) as zipf:
    for root, _, files in os.walk(temp_folder):
        for file in files:
            file_path = os.path.join(root, file)
            zipf.write(file_path, os.path.relpath(file_path, temp_folder))

# Geçici klasörü temizler
shutil.rmtree(temp_folder)

olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)
newmail.Subject = 'Todays advices as of ' + datetime.now().strftime('%#d %b %Y')
newmail.To = 'bemretuyluoglu@gmail.com'
newmail.CC = 'bemretuyluoglu@gmail.com'
newmail.Body = 'todays advice is attached.'

# Arşivlenmiş zip dosyasını ekle
newmail.Attachments.Add(zip_file_path)

# E-postayı görüntülemek üzere
newmail.Display()
#newmail.Send()
