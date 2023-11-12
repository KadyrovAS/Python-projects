import subprocess
from sys import argv

def prt_code(code):
    ln = ""
    for simbol in code:
        ln = f"{ln} {ord(simbol)}"
    print(ln)
def prt(file_in):
    fn = open(file_in, "r")
    lines = fn.readlines()

    for line in lines:
        ln = ""

        for simbol in line:
            ln = f"{ln} {ord(simbol)}"
        print(ln)


# print(argv)

file_in = "002.txt"
file_out = "002res.txt"
code = chr(1) * 5

subprocess.Popen(f"encoder-windows-amd64.exe {code} {file_in} {file_out}")


print("Исходный файл")
prt(file_in)
print("результирующий файл")
prt(file_out)
print("Код")
prt_code(code)