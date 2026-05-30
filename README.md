# ListShuffle-DupCleaner 🛠️

Aplikasi desktop berbasis Python dan Tkinter yang ringan dan praktis untuk mengelola daftar teks (*list*). Aplikasi ini menyediakan fitur untuk membagi daftar teks ke dalam file Excel (dengan opsi pengacakan otomatis) serta membersihkan item duplikat atau menyaring data menggunakan daftar pembanding.

## 🚀 Fitur Utama

### 1. Pembagi List (List Splitter & Shuffler)
* **Penghitung Real-time:** Menghitung total item yang dimasukkan secara langsung saat Anda mengetik atau menempelkan teks.
* **Dua Mode Pembagian Kolom:**
  * **Jumlah Kolom:** Membagi seluruh item secara merata ke dalam jumlah kolom yang Anda tentukan.
  * **Item per Kolom:** Menentukan batasan jumlah baris/item yang persis untuk setiap kolomnya.
* **Opsi Pengacakan (Shuffle):** Mengacak urutan baris secara acak sebelum memproses data.
* **Ekspor ke Excel:** Membuat dan memformat spreadsheet `.xlsx` secara dinamis menggunakan pustaka `openpyxl`.

### 2. Duplikat List (Duplicate Cleaner & Comparator)
* **Pembersih Duplikat:** Menemukan dan menghapus data yang ganda secara instan sekaligus mengurutkannya secara alfabetis (A-Z).
* **Komparator / Pembanding Data:** Membandingkan "Data Utama" dengan "Data Pembanding", lalu menghapus item di Data Utama jika item tersebut ditemukan di Data Pembanding.
* **Log Visual:** Memisahkan tampilan visual antara hasil bersih (data unik) dan daftar item yang berhasil dihapus/dibersihkan.

---

## 🛠️ Persyaratan Sistem

* Python 3.x
* Pustaka `openpyxl` (untuk manipulasi ekspor file Excel)

---

## ⚙️ Cara Instalasi & Penggunaan

1. *Clone* repositori ini ke komputer Anda:
```bash
   git clone [https://github.com/qwertyaqu-prog/ListShuffle-DupCleaner.git](https://github.com/qwertyaqu-prog/ListShuffle-DupCleaner.git)
```

2. Masuk ke dalam folder proyek:
```Bash
   cd ListShuffle-DupCleaner
```
3. Instal dependensi atau library yang diperlukan:
```Bash
   pip install -r requirements.txt
```
4. Jalankan aplikasinya:
```Bash
   python ListShuffle-DupCleaner.py
```
## 📄 Lisensi
Proyek ini dilindungi di bawah Lisensi MIT - silakan lihat file LICENSE untuk detail lebih lanjut.
