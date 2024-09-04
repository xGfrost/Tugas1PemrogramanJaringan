def f(x):
    if x != 0:
        return 2 * (x ** 3) + 2 * x + 15 / x
    else:
        return "Tidak terdefinisi (x tidak boleh 0)"

x = int(input("Masukkan nilai x: "))
hasil = f(x)
print(f"f(x) = {hasil}")
