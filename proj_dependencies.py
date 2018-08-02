
def taskDependencyRange(start, end):
    arr = []
    for i in range(start, end+1):
        arr.append(str(i))
    print(','.join(arr))


taskDependencyRange(148,156)