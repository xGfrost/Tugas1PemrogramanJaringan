def seleksi_kriteria(nilai):
    if nilai >= 81:
        return "A", "Istimewa"
    elif nilai >= 71:
        return "AB", "Sangat baik"
    elif nilai >= 66:
        return "B", "Baik"
    elif nilai >= 61:
        return "BC", "Cukup baik"
    elif nilai >= 56:
        return "C", "Cukup"
    elif nilai >= 41:
        return "D", "Kurang"
    else:
        return "E", "Sangat kurang"

nilai = int(input("Masukkan nilai: "))
kriteria, kategori = seleksi_kriteria(nilai)
print(f"Kriteria nilai: {kriteria}, Kategori: {kategori}")
