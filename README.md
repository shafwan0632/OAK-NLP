import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from transformers import pipeline
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd

# Inisialisasi sentiment analyzer
sentiment_analyzer = pipeline("text-classification",
                              model="ayameRushia/bert-base-indonesian-1.5G-sentiment-analysis-smsa")


# Fungsi untuk scroll
def scroll(height):
    driver.execute_script(f"window.scrollTo(0, {height});")


# Inisialisasi WebDriver
option = webdriver.ChromeOptions()
servicePath = Service('D:/new life/pythonProject/test/direct1/packageUtama/tkinter/chromedriver.exe')
driver = webdriver.Chrome(service=servicePath, options=option)

# Masuk ke halaman dan mencari produk
driver.get("https://www.tokopedia.com/")
driver.maximize_window()
time.sleep(3)
searchbar = driver.find_element(By.CLASS_NAME, "css-3017qm")
searchbar.send_keys("Mystery Box", Keys.ENTER)
time.sleep(5)
scroll(300)

product = driver.find_element(By.XPATH,
                              "//span[text()='Mistery Box Dapat hp - unboxing misteri box misterius / misteri dapat']")
product.click()
time.sleep(5)

scroll(3600)
time.sleep(5)

# Mengumpulkan komentar
all_comments = []
page = 1
while True:
    try:
        comments = driver.find_elements(By.XPATH, "//span[@data-testid='lblItemUlasan']")
        for comment in comments:
            all_comments.append(comment.text)

        next_button = driver.find_element(By.XPATH, f'//button[@aria-label="Laman {page + 1}"]')
        next_button.click()
        time.sleep(3)
        page += 1
    except:
        break

# Analisis sentimen untuk setiap komentar
sentiments = []
scores = []

for comment in all_comments:
    sentiment = sentiment_analyzer(comment)
    sentiments.append(sentiment[0]['label'])
    scores.append(sentiment[0]['score'])

# Membuat DataFrame untuk visualisasi
df = pd.DataFrame({
    'Comment': all_comments,
    'Sentiment': sentiments,
    'Score': scores
})

# Membuat window tkinter
window = tk.Tk()
window.title("Analisis Sentimen")
window.geometry("1000x800")

# Mengatur font global menjadi Comic Sans
window.option_add("*Font", "Comic Sans MS 10")

# Frame untuk plot dan tabel
top_frame = tk.Frame(window)
top_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

bottom_frame = tk.Frame(window)
bottom_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

# Membuat plot seaborn
plt.figure(figsize=(8, 4))
sns.countplot(x='Sentiment', data=df, palette='Set2')
plt.title('Distribusi Sentimen')

# Mengonversi plot ke canvas tkinter
canvas_plot = FigureCanvasTkAgg(plt.gcf(), master=top_frame)
canvas_plot.draw()
canvas_plot.get_tk_widget().pack()

# Menambahkan tabel komentar menyerupai Excel
tree = ttk.Treeview(bottom_frame, columns=("No", "Comment", "Sentiment", "Confidence"), show="headings")
tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Menambahkan scrollbar pada tabel
scrollbar = ttk.Scrollbar(bottom_frame, orient="vertical", command=tree.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
tree.configure(yscrollcommand=scrollbar.set)

# Menambahkan header pada tabel
tree.heading("No", text="No")
tree.heading("Comment", text="Comment")
tree.heading("Sentiment", text="Sentiment")
tree.heading("Confidence", text="Confidence")

# Mengatur lebar kolom
tree.column("No", width=50, anchor=tk.CENTER)
tree.column("Comment", width=500)
tree.column("Sentiment", width=150, anchor=tk.CENTER)
tree.column("Confidence", width=150, anchor=tk.CENTER)

# Menambahkan data ke tabel dengan nomor urut
for idx, (comment, sentiment, score) in enumerate(zip(all_comments, sentiments, scores), start=1):
    tree.insert("", tk.END, values=(idx, comment, sentiment, f"{score:.2f}"))

# Mengatur tampilan tabel menyerupai Excel
style = ttk.Style()
style.configure("Treeview.Heading", font=("Comic Sans MS", 10, "bold"))
style.configure("Treeview", rowheight=25, font=("Comic Sans MS", 10),
                borderwidth=1, relief="solid")
style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])  # Hilangkan garis grid eksternal

# Jalankan aplikasi
window.mainloop()
