import random, time
import openpyxl
import numpy as  np


class Gacha():
    def __init__(self, rate):
        self.rate = [rate, 100 - rate]
        self.gacha_list = [1, 0]
    
    # continuous 連ガチャをgacha_count 回引いた時の確立
    def virtual_gacha(self, gacha_count, continuous):
        num = 0
        for _ in range(gacha_count):
            result = random.choices(self.gacha_list, k=continuous, weights=self.rate)
            num += result.count(1)
        
        return num
    
    def gacha_rate(self, gacha_count, continuous):
        win = self.virtual_gacha(gacha_count, continuous)
        rate = win / (gacha_count * continuous) * 100
        
        return rate
    
    def first_get_count(self):
        get_count = 0
        result = 0
        
        while(result != 1):
            result = random.choices(self.gacha_list, weights=self.rate)[0]
            get_count += 1
        
        return get_count
    
    def average_first_count(self, gacha_count):
        # 範囲の個数
        label_size = 15
        # 範囲
        range_size = 20
        
        ranges = [range_size * (x + 1) for x in range(label_size)]
        labels = [f"{range_size * x + 1} ~ {range_size * x + range_size}" for x in range(label_size)] + [f"{(label_size) * range_size + 1} ~ "]
        gacha_time = {label: 0 for label in labels}
        
        max = 0
        min = label_size * range_size
        average = 0
        count_list = []
        for i in range(gacha_count):
            count = self.first_get_count()
            count_list.append(count)
            average += count
            if i % 100 == 0:
                print(f"\r{int(i / (gacha_count / 100))}% 完了", end="")
            # 最大値
            if count > max:
                max = count
            # 最小値
            if count < min:
                min = count
            # 辞書
            for i, limit in enumerate(ranges):
                if count <= limit:
                    gacha_time[labels[i]] += 1
                    break
            else:
                gacha_time[labels[-1]] += 1
        median = np.median(count_list)
        average /= gacha_count
        print("\r100% 完了\n")
        # 試行回数の辞書、最大試行回数、最小試行回数、中央値、平均値
        return gacha_time, max, min, median, average


def export_to_excel(result_list, file_name="gacha_results.xlsx"):
    # 新しいExcelワークブックを作成
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gacha Results"

    # データを書き込み
    ws.append(["Virtual Gacha Result"])
    # [0]
    ws.append(["確率", "", result_list[0]])
    # [5]
    ws.append(["平均試行回数", "", f"{result_list[5]}連"])
    # [4]
    ws.append(["中央値", "", f"{result_list[4]}連"])
    # [3]
    ws.append(["最小試行回数", "", f"{result_list[3]}連"])
    # [2]
    ws.append(["最大試行回数", "", f"{result_list[2]}連"])
    ws.append([])  # 空行

    ws.append(["Range", "Count", "rate", "comulative"])
    i = 0
    # [1], [6], [7]
    for k, v in result_list[1].items():
        ws.append([k, v, result_list[6][i], result_list[7][i]])
        i += 1

    # ファイル保存
    wb.save(file_name)
    print(f"\nResults saved to {file_name}")


def main():
    # 仮想ガチャ、各回の0, 1 生成回数、x 連ガチャ
    continuous = 10
    # 試行回数
    gacha_count = 10000
    # 確率 x %
    rate = 3
    
    gacha = Gacha(rate)

    start = time.time()

    rate_result = gacha.gacha_rate(gacha_count, continuous)
    average_result_dict, max, min, median, average_rate = gacha.average_first_count(gacha_count)

    end = time.time()
    
    print(f"当たり確率 {rate}% のガチャの試行")
    total_rate = 0
    cumulative_rate_list = []
    each_rate_list = []
    for k, v in average_result_dict.items():
        x = v /gacha_count * 100
        total_rate += x
        cumulative_rate_list.append(total_rate)
        each_rate_list.append(x)
        print(f"{k:^10}連: {v:>8}回: {total_rate:>6.2f}%")
    print(f"\n試行時の確率: {rate_result:>6.2f}%")
    print(f"試行回数    : {gacha_count:>6}回")
    print(f"中央値      : {median:>6}回")
    print(f"平均値      : {average_rate:>6.2f}回")
    print(f"最大試行回数: {max:>6}回")
    print(f"最小試行回数: {min:>6}回")
    
    # 試行時の確率、結果の辞書、最大試行回数、最小、中央値、平均、各確率、累積確率
    to_excel_list = [
        rate_result,
        average_result_dict,
        max,
        min,
        median,
        average_rate,
        each_rate_list,
        cumulative_rate_list,
        ]
        
    export_to_excel(to_excel_list)
    
    
    time_diff = end - start
    print(f"\n{time_diff}s")


if __name__ == "__main__":
    main()
