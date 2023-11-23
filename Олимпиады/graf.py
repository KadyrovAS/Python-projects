def set_map(i, j, sum, n):
    global N, M, K
    global ar, tele, tbl

    n += 1
    sum += tbl[i][j]
    if sum / n > ar[i][j][0] / ar[i][j][1]:
        ar[i][j][0] = sum
        ar[i][j][1] = n
    else:
        return

    if i == 0 and j == 0:
        return

    if i > 0:
        set_map(i - 1, j, sum, n)


    if j > 0:
        set_map(i, j - 1, sum, n)


N, M, K = map(int, input().split())
ar = [[[0, 0] for _ in range(M)] for _ in range(N)]
tele = [list(map(int, input().split())) for _ in range(K)] #координаты телепортов
tbl = [list(map(int, input().split())) for _ in range(N)]
set_map(N - 1, M - 1, 0, 0)

# coord = []
# max = 0
# for t in tele:
#     n = t[0] - 1
#     m = t[1] - 1
#     if ar[n][m][0] > max:
#         max = ar[n][m][0]
#         coord = [n, m]
# sum = ar[0][0][0] * ar[0][0][1]
# n = ar[0][0][1]

for line in ar:
    print(line)