a = int(input())
args = []
n = a
while n != 1:
    for i in range(2, n + 1):
        if n % i == 0:
            args.append(i)
            n //= i
            break

print(args)