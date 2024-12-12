import random, time
import openpyxl
import numpy as  np
from array import array


class Gacha():
    def __init__(self, rate):
        self.rate = [rate, 100 - rate]
        self.gacha_list = [1, 0]
    
    # 仮想ガチャ、continuous 連をgacha_count回行う
    def virtual_gacha(self, gacha_count, continuous=1):
        num = 0
        for _ in range(gacha_count):
            result = random.choices(self.gacha_list, k=continuous, weights=self.rate)
            num += result.count(1)
        
        return num
    
    # ガチャの確率を求める
    def gacha_rate(self, gacha_count):
        win = self.virtual_gacha(gacha_count)
        rate = win / gacha_count * 100
        
        return rate
    
    # 初めてあたりを引けるまで行うガチャ
    def first_get_count(self):
        get_count = 0
        result = 0
        
        while(result != 1):
            result = random.choices(self.gacha_list, weights=self.rate)[0]
            get_count += 1
        
        return get_count
    
    # 初めてあたりを引くことをgacha_cont 回行う
    def rate_first_get(self, gacha_count):
        # 範囲の個数
        label_size = 15
        # 範囲
        range_size = 20
        
        ranges = array("i", [range_size * (x + 1) for x in range(label_size)])
        labels = [f"{range_size * x + 1} ~ {range_size * x + range_size}" for x in range(label_size)] + [f"{(label_size) * range_size + 1} ~ "]
        gacha_time = {label: 0 for label in labels}
        
        max = 0
        min = label_size * range_size
        average = 0
        count_list = array("i", [])
        print("\n初めて獲得するまで / Roll a gacha until getting item first")
        for i in range(gacha_count):
            count = self.first_get_count()
            count_list.append(count)
            average += count
            if i % 100 == 0:
                print(f"\r{int(i / (gacha_count / 100))}% completed", end="")
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
        print("\r100% completed      ")
        # 試行回数の辞書、最大試行回数、最小試行回数、中央値、平均値
        return gacha_time, max, min, median, average
    
    # 課金額以内で当たりを引ける確率
    def budget_gacha(self, gacha_count, number_of_gacha):
        get = 0
        get = array("i", [])
        print("\n課金 / Budget Gacha")
        # number_of_gacha 連をgacha_count 回引く
        for i in range(gacha_count):
            win = self.virtual_gacha(gacha_count=1, continuous=number_of_gacha)
            get.append(win)
            print(f"\r{int(i  / (gacha_count / 100))}% completed", end="")

        print("\r100% completed    \n")
        return get


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
    # 試行回数 / Number of trials
    gacha_count = 10000
    # 確率 / Rate (rate %)
    rate = 0.1
    # 課金額 / Budget (budget yen)
    budget = 20000
    # 1連の単価 / Cost of one roll 
    cost = 300
    # ガチャを回せる回数 / Number of time rolling gacha within budget
    number_of_gacha = int(budget / cost)
    
    gacha = Gacha(rate)

    start = time.time()

    rate_result = gacha.gacha_rate(gacha_count)
    rate_result_dict, max, min, median, average_rate = gacha.rate_first_get(gacha_count)
    budget_gacha_rate_list = gacha.budget_gacha(gacha_count, number_of_gacha)

    end = time.time()
    
    print(f"{rate}% の当たり確率で最初にあたりを引くまでに回した回数 / Number of times until winning (gacha with a {rate}% chance of winning)")
    total_rate = 0
    cumulative_rate_list = array("f", [])
    each_rate_list = array("f", [])
    for k, v in rate_result_dict.items():
        each_rate = v / gacha_count * 100
        total_rate += each_rate
        cumulative_rate_list.append(total_rate)
        each_rate_list.append(each_rate)
        print(f"{k:^10} 連 / Rolls: {v:>8} 回 / times: {each_rate:>6.2f} %: {total_rate:>6.2f} %")
    print(f"\n試行時の確率 / Rate                    : {rate_result:>6.2f} %")
    print(f"試行回数 / Number of trials            : {gacha_count:>6} 回 / times")
    print(f"中央値 / Median                        : {median:>6} 回 / times")
    print(f"平均値 / Average                       : {average_rate:>6.2f} 回 / times")
    print(f"最大試行回数 / Maximum number of trials: {max:>6} 回 / times")
    print(f"最小試行回数 / Minimum number of trials: {min:>6} 回 / times\n")
    
    i = 0
    total = 0
    print(f"\n{budget}円課金した場合（{number_of_gacha}連) / When charging {budget}yen ({number_of_gacha} rolls)")
    print("獲得数:      獲得回数  : パーセンテージ/ Number of wins: Number of times: Percentage")
    while i < 3:
        get = budget_gacha_rate_list.count(i)
        get_rate = get / len(budget_gacha_rate_list) * 100
        print(f"{i:^6}: {get:>6} 回/times: {get_rate:>6.2f} %")
        total += get
        i += 1
    get = len(budget_gacha_rate_list) - total
    get_rate = get / len(budget_gacha_rate_list) * 100
    print(f"  {i} ~ : {get:>6} 回/times: {get_rate:>6.2f} %")
        
    
    # 試行時の確率、結果の辞書、最大試行回数、最小、中央値、平均、各確率、累積確率
    to_excel_list = [
        rate_result,
        rate_result_dict,
        max,
        min,
        median,
        average_rate,
        each_rate_list,
        cumulative_rate_list,
        ]
    
    # Excelに書き込み 
    # export_to_excel(to_excel_list)
    
    
    time_diff = end - start
    print(f"\n{time_diff:.4f} seconds")


if __name__ == "__main__":
    main()
