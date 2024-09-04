import statistics as stats

def hitung_rata2(n1, n2, n3):
    return (n1 + n2 + n3) / 3

def hitung_statistik(data):
    rata2 = stats.mean(data)
    median = stats.median(data)
    try:
        modus = stats.mode(data)
    except:
        modus = "Tidak ada modus"
    return rata2, median, modus

# Contoh untuk rata-rata dari 3 bilangan
n1, n2, n3 = map(int, input("Masukkan 3 bilangan bulat: ").split())
rata2 = hitung_rata2(n1, n2, n3)
print(f"Rata-rata dari 3 bilangan: {rata2}")

# Contoh untuk menghitung rata2, median, dan modus dari 10 data
data = list(map(int, input("Masukkan 10 data bilangan: ").split()))
rata2, median, modus = hitung_statistik(data)
print(f"Rata-rata: {rata2}, Median: {median}, Modus: {modus}")
