def s(n, doc):
    summingtime = doc[0][0]
    wait = []
    total_wait = 0
    emptyTime = 0
    for i in range(n):
        emptyTime = max(0, doc[i][0] - summingtime)
        summingtime += doc[i][1] + emptyTime
        waitTime = summingtime - doc[i][0]
        total_wait += waitTime
    return total_wait


s1 = s(3, [[3,2], [5, 4], [7,5]])
print(s1)

# 2 2
# 5 5