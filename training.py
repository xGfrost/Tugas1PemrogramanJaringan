import openpyxl
import pandas as pd
import os
import time

FILE_EXCEL = "data_mahasiswa.xlsx"


# Fungsi untuk menghitung nilai huruf, indeks, dan kriteria
def hitung_indeks(nilai):
    if nilai >= 88:
        return 'A', 4.00, 'Istimewa'
    elif 77 <= nilai < 88:
        return 'B', 3.00, 'Baik'
    elif 60 <= nilai < 77:
        return 'C', 2.00, 'Cukup'
    elif 45 <= nilai < 60:
        return 'D', 1.00, 'Sedang'
    else:
        return 'E', 0.00, 'Kurang'

# Fungsi untuk menambahkan mahasiswa
def tambah_mahasiswa(mahasiswa_list):
    nama = input("Masukkan nama mahasiswa: ")
    
    while True:
        try:
            nilai = float(input("Masukkan nilai mahasiswa (0-100): "))
            if 0 <= nilai <= 100:
                break
            else:
                print("Nilai harus berada di antara 0 dan 100. Coba lagi.")
        except ValueError:
            print("Nilai harus berupa angka.")
    
    nilai_huruf, indeks, kriteria = hitung_indeks(nilai)
    mahasiswa_list.append([nama, nilai, nilai_huruf, indeks, kriteria])
    
    simpan_ke_excel(mahasiswa_list)
    print(f"Mahasiswa {nama} berhasil ditambahkan.")
    preview_excel()


# Fungsi untuk membaca semua data mahasiswa
def baca_semua_mahasiswa(mahasiswa_list):
    if mahasiswa_list:
        for mahasiswa in mahasiswa_list:
            print(f"Nama: {mahasiswa[0]}, Nilai: {mahasiswa[1]}, Huruf: {mahasiswa[2]}, Indeks: {mahasiswa[3]}, Kriteria: {mahasiswa[4]}")
    else:
        print("Tidak ada data mahasiswa.")
    preview_excel()


# Fungsi untuk membaca data mahasiswa berdasarkan nama
def baca_mahasiswa(mahasiswa_list):
    nama = input("Masukkan nama mahasiswa yang ingin dibaca: ")
    
    for mahasiswa in mahasiswa_list:
        if mahasiswa[0] == nama:
            print(f"Nama: {mahasiswa[0]}, Nilai: {mahasiswa[1]}, Huruf: {mahasiswa[2]}, Indeks: {mahasiswa[3]}, Kriteria: {mahasiswa[4]}")
            return
    print(f"Mahasiswa dengan nama {nama} tidak ditemukan.")


# Fungsi untuk memperbarui nilai mahasiswa
def update_mahasiswa(mahasiswa_list):
    nama = input("Masukkan nama mahasiswa yang ingin diupdate: ")
    
    for mahasiswa in mahasiswa_list:
        if mahasiswa[0] == nama:
            while True:
                try:
                    nilai_baru = float(input("Masukkan nilai baru mahasiswa (0-100): "))
                    if 0 <= nilai_baru <= 100:
                        break
                    else:
                        print("Nilai harus berada di antara 0 dan 100. Coba lagi.")
                except ValueError:
                    print("Nilai harus berupa angka.")
            
            nilai_huruf, indeks, kriteria = hitung_indeks(nilai_baru)
            mahasiswa[1], mahasiswa[2], mahasiswa[3], mahasiswa[4] = nilai_baru, nilai_huruf, indeks, kriteria
            
            simpan_ke_excel(mahasiswa_list)
            print(f"Nilai mahasiswa {nama} berhasil diperbarui.")
            preview_excel()
            return
    print(f"Mahasiswa dengan nama {nama} tidak ditemukan.")


# Fungsi untuk menghapus mahasiswa
def hapus_mahasiswa(mahasiswa_list):
    nama = input("Masukkan nama mahasiswa yang ingin dihapus: ")
    
    for mahasiswa in mahasiswa_list:
        if mahasiswa[0] == nama:
            mahasiswa_list.remove(mahasiswa)
            simpan_ke_excel(mahasiswa_list)
            print(f"Mahasiswa {nama} berhasil dihapus.")
            preview_excel()
            return
    print(f"Mahasiswa dengan nama {nama} tidak ditemukan.")


# Fungsi untuk menyimpan data ke Excel
def simpan_ke_excel(mahasiswa_list):
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Mahasiswa"
        ws.append(["Nama", "Nilai", "Huruf", "Indeks", "Kriteria"])
        
        for mahasiswa in mahasiswa_list:
            ws.append(mahasiswa)
        
        wb.save(FILE_EXCEL)
    except PermissionError:
        print(f"PermissionError: Tidak dapat menyimpan file {FILE_EXCEL}. Pastikan file tidak sedang dibuka di aplikasi lain.")
    except Exception as e:
        print(f"Terjadi kesalahan saat menyimpan file Excel: {e}")


# Fungsi untuk menampilkan preview data dari Excel
def preview_excel():
    try:
        df = pd.read_excel(FILE_EXCEL)
        print("\nLive Preview dari File Excel:")
        print(df)
        os.system(f'start excel {FILE_EXCEL}')  # Windows
        time.sleep(2)  # Delay to allow Excel to open
    except FileNotFoundError:
        print("File Excel belum dibuat.")
    except PermissionError:
        print(f"PermissionError: Tidak dapat membuka file {FILE_EXCEL}. Pastikan file tidak sedang dibuka di aplikasi lain.")
    except Exception as e:
        print(f"Terjadi kesalahan saat membuka file Excel: {e}")


# Fungsi utama untuk menampilkan menu
def menu():
    mahasiswa_list = []
    while True:
        print("\nMenu:")
        print("1. Tambah Mahasiswa")
        print("2. Baca Semua Mahasiswa")
        print("3. Baca Mahasiswa")
        print("4. Update Mahasiswa")
        print("5. Hapus Mahasiswa")
        print("6. Keluar")
        
        pilihan = input("Pilih menu (1-6): ")
        
        if pilihan == '1':
            tambah_mahasiswa(mahasiswa_list)
        elif pilihan == '2':
            baca_semua_mahasiswa(mahasiswa_list)
        elif pilihan == '3':
            baca_mahasiswa(mahasiswa_list)
        elif pilihan == '4':
            update_mahasiswa(mahasiswa_list)
        elif pilihan == '5':
            hapus_mahasiswa(mahasiswa_list)
        elif pilihan == '6':
            print("Keluar dari program.")
            break
        else:
            print("Pilihan tidak valid. Silakan pilih menu yang benar.")


if __name__ == "__main__":
    menu()