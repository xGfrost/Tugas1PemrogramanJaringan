import numpy as np

def penjumlahan_matriks(matriks1, matriks2):
    return np.add(matriks1, matriks2)

matriks1 = np.array([[1, 2], [3, 4]])
matriks2 = np.array([[5, 6], [7, 8]])

hasil = penjumlahan_matriks(matriks1, matriks2)
print("Hasil penjumlahan matriks:")
print(hasil)
