def fact(n):
    return 1 if n == 0 else n * fact(n - 1)
w = "превысокомногорассмотрительствующий"
n = 5

print(f"{n}! = {fact(n)}")