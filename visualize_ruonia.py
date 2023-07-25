import matplotlib.pyplot as plt

def visualize_ruonia(data, count):
    plt.close("all")

    font = {'family' : 'monospace',
            'size'   : 8}
    plt.rc('font', **font)
    plt.grid()
    plt.xticks(rotation=90)

    colors = [(0,0,1), (1,0,0), (0,1,0)]
    for idx, dataItem in enumerate(data):
        plt.plot(*zip(*dataItem), color=colors[idx])

    # Настройки шага по оси X
    plt.gca().margins(x=0)
    plt.gcf().canvas.draw()
    tl = plt.gca().get_xticklabels()
    maxsize = max([t.get_window_extent().width for t in tl])
    m = 0.2 # inch margin
    s = maxsize/plt.gcf().dpi*count+2*m
    margin = m/plt.gcf().get_size_inches()[0]

    plt.gcf().subplots_adjust(left=margin, right=1.-margin)
    plt.gcf().set_size_inches(s, plt.gcf().get_size_inches()[1])
    # Конец настройки шага по оси X

    plt.show()