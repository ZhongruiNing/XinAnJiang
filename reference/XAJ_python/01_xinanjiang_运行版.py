# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import spotpy
from pack.best_para_get import best_paras_from_name


def EvaporationCalc(WUM, WLM, WDM, W, WU, WL, WD, Ep, P, C, WMM, WM, B):
    """
    计算土壤水分蒸发过程中的各种参数变化。
    参数:
    WUM: 最大允许上层土壤含水量/蓄水能力
    WLM: 最大允许下层土壤含水量/蓄水能力
    WDM: 最大允许深层土壤含水量/蓄水能力
    W: 初始土壤含水量
    WU: 上层土壤含水量
    WL: 下层土壤含水量
    WD: 深层土壤含水量
    Ep: 时段潜在蒸发量
    P: 时段降水量
    C: 深层蒸散发系数
    WMM: 流域土壤水分存储能力/单点最大蓄水量
    WM: 土壤水分饱和度/平均张力蓄水容量
    B: 蓄水容量-面积分配曲线的指数

    返回:
    E: 总蒸发量
    W_next: 下一时刻的土壤含水量
    WL_next: 下一时刻的中间层土壤含水量
    WD_next: 下一时刻的下层土壤含水量
    WU_next: 下一时刻的上层土壤含水量
    R: 产流量
    PE: 净雨量
    """
    n = len(Ep)
    # 初始化时段蒸发量、渗透量和产流量数组
    Eu = np.zeros(n)
    EL = np.zeros(n)
    ED = np.zeros(n)
    E = np.zeros(n)
    PE = np.zeros(n)
    W_next = np.zeros(n + 1)
    WU_next = np.zeros(n + 1)
    WL_next = np.zeros(n + 1)
    WD_next = np.zeros(n + 1)
    R = np.zeros(n)

    for i in range(n):
        # i代表时段
        # 第一小步: 计算各层土壤的蒸发量
        # 参考教材《水文预报》（第五版）P146-147页
        if P[i] + WU[i] >= Ep[i]:
            Eu[i] = Ep[i]
        else:
            Eu[i] = P[i] + WU[i]
            # 根据中间层土壤含水量计算蒸发量
            if WL[i] >= C * WLM:
                EL[i] = (Ep[i] - Eu[i]) * WL[i] / WLM
            elif WL[i] >= C * (Ep[i] - Eu[i]):
                EL[i] = C * (Ep[i] - Eu[i])
            else:
                EL[i] = WL[i]
                ED[i] = C * (Ep[i] - Eu[i]) - EL[i]
        E[i] = Eu[i] + EL[i] + ED[i]

        # 计算下渗量和产流量
        # 第二小步: 利用蓄满产流模型进行产流计算
        # 参考教材《水文预报》（第五版）P148页
        PE[i] = P[i] - E[i]
        A = WMM * (1 - (1 - W[i] / WM) ** (1 / (1 + B)))
        if PE[i] <= 0:
            R[i] = 0
        elif PE[i] + A < WMM:
            # 局部产流
            R[i] = PE[i] - WM + W[i] + WM * (1 - (PE[i] + A) / WMM) ** (1 + B)
        else:
            # 全流域产流
            R[i] = PE[i] - WM + W[i]

        # 更新各层土壤含水量
        # 计算各层下一时刻初土壤含水量
        if i < n:
            # 确保上层土壤含水量在合理范围内
            WU_next[i + 1] = P[i] + WU[i] - Eu[i] - R[i]
            temp1 = 0  # 临时变量
            temp2 = 0
            if WU_next[i + 1] > WUM:
                R1 = WU_next[i + 1] - WUM
                WU_next[i + 1] = WUM
            elif WU_next[i + 1] <= 0:
                temp1 = abs(WU_next[i + 1])
                WU_next[i + 1] = 0
                R1 = 0
            else:
                R1 = 0

            # 确保中间层土壤含水量在合理范围内
            WL_next[i + 1] = WL[i] + R1 - EL[i] - temp1
            if WL_next[i + 1] > WLM:
                R2 = WL_next[i + 1] - WLM
                WL_next[i + 1] = WLM
            elif WL_next[i + 1] <= 0:
                temp2 = abs(WL_next[i + 1])
                WL_next[i + 1] = 0
                R2 = 0
            else:
                R2 = 0

            # 确保下层土壤含水量在合理范围内
            WD_next[i + 1] = WD[i] + R2 - ED[i] - temp2
            if WD_next[i + 1] > WDM:
                R3 = WD_next[i + 1] - WDM  # 仅为了与之前代码风格一致
                WD_next[i + 1] = WDM

            W_next[i + 1] = WL_next[i + 1] + WU_next[i + 1] + WD_next[i + 1]

    return E, W_next, WL_next, WD_next, WU_next, R, PE


def XinAnJiang(WM, WUM, WLM, B, C, SM, EX, KG, KI, KKG, KKSS, CS, FE, IMP, P, Ep):
    """
    新安江模型的核心计算函数，用于模拟流域的水文过程。
    参数:
    WM: 土壤水分饱和度/平均张力蓄水容量
    WUM: 最大允许上层土壤含水量/蓄水能力
    WLM: 最大允许下层土壤含水量/蓄水能力
    B: 河槽蓄水能力参数/张力水
    C: 深层蒸散发系数
    SM: 表层土壤最大蓄水量
    EX: 土壤蓄水饱和度指数/表层自由水蓄水容量的方次
    KG: 地表径流系数
    KI: 土壤深层出流系数
    KKG: 地下水出流系数
    KKSS: 河槽床面渗漏系数
    CS: 消退系数
    FE: 初始土壤含水量比例/产流面积
    IMP: 不透水面积比例
    P: 降水量序列
    Q: 径流量序列
    Ep: 蒸发能力

    返回值:
    无直接返回值，但通过全局变量或外部变量影响其他部分的计算。
    """
    # 定义一个长时间零值（空值）序列，以覆盖降雨数据历时
    n = len(P)  #
    # 初始化蓄水量数组
    EP = Ep  # 流域蒸发能力
    WU = np.zeros(n + 1)
    WL = np.zeros(n + 1)
    WD = np.zeros(n + 1)
    # 根据初始含水量FE，计算各蓄水层的初始值
    WU[0] = WUM * FE
    WL[0] = WLM * FE
    WDM = WM - WUM - WLM
    WD[0] = WDM * FE
    # 计算调整后的时段最大蓄水量
    WMM = WM * (1 + B)
    # 计算每时段的总蓄水量
    W = WL + WD + WU

    # 调用蒸发计算函数，更新蓄水量和径流量等，进行产流计算
    E, W, _, WD, WU, R, PE = EvaporationCalc(WUM, WLM, WDM, W, WU, WL, WD, EP, P, C, WMM, WM, B)

    # 汇流计算部分
    # 重塑径流量数组，以匹配其他数组的维度
    R = R.reshape(-1)
    # 计算调整后的土壤最大蓄水量
    SMM = SM * (1 + EX)
    # 初始化存储相关变量的数组
    S = np.zeros(n)
    RS = np.zeros(n)
    RG = np.zeros(n)
    RI = np.zeros(n)
    S[0] = FE * SM
    FR = np.zeros(n)

    # 遍历每一天，计算蓄水和径流情况
    for i in range(n-1):
        # 计算可更新的土壤蓄水量
        # 参考教材《水文预报》（第五版）P150页 公式5.20
        AU = SMM * (1 - (1 - S[i] / SM) ** (1 / (1 + EX)))

        # 根据蓄水条件，更新径流和蓄水情况
        # 如果净雨量为0，则通过流域已有蓄水量产流，地表部分不产流
        if PE[i] <= 0:
            if np.abs(W[i] - WM) < 0.001:
                FR[i] = 0
            else:
                FR[i] = 1 - (1 - W[i] / WM) ** (B / (1 + B))

            RS[i] = 0
            RI[i] = S[i] * KI * FR[i]
            RG[i] = S[i] * KG * FR[i]
            S[i + 1] = (1 - KI - KG) * S[i]
        # 如果净雨量大于0，则根据蓄满产流计算地表产流量
        else:
            if PE[i] + AU < SMM:
                FR[i] = R[i] / PE[i]
                RS[i] = FR[i] * (PE[i] - SM + S[i] + SM * (1 - (PE[i] + AU) / SMM) ** (EX + 1))
                RI[i] = KI * FR[i] * (SM - SM * (1 - (PE[i] + AU) / SMM) ** (EX + 1))
                RG[i] = KG * FR[i] * (SM - SM * (1 - (PE[i] + AU) / SMM) ** (EX + 1))
                S[i + 1] = (1 - KI - KG) * (SM - SM * (1 - (PE[i] + AU) / SMM) ** (1 + EX))
            else:
                FR[i] = R[i] / PE[i]
                RS[i] = (PE[i] - SM + S[i]) * FR[i]
                RI[i] = SM * KI * FR[i]
                RG[i] = SM * KG * FR[i]
                S[i + 1] = (1 - KI - KG) * SM

    # 根据不透水面积比例系数，调整径流成分
    RS = RS * (1 - IMP) + P[:n] * IMP
    RG = RG * (1 - IMP)
    RI = RI * (1 - IMP)
    # 计算总径流量
    Rt = RS + RG + RI  # 水源划分后总径流
    RS[np.isnan(RS)] = 0
    RG[np.isnan(RG)] = 0
    RI[np.isnan(RI)] = 0

    n1 = len(P)
    F = basin_area
    dt = time_scale
    U = F / (3.6 * dt)
    QRS = RS * U
    QS = np.zeros(n1)
    QI = np.zeros(n1)
    QG = np.zeros(n1)

    # 利用线性滞后算法进行汇流计算
    QS[0] = (1 - CS) * RS[0] * U
    QI[0] = (1 - KKSS) * RI[0] * U
    QG[0] = (1 - KKG) * RG[0] * U
    for i in range(1, n1-2):
        QS[i] = KKSS * QS[i - 1] + (1 - CS) * RS[i] * U
        QI[i] = KKSS * QI[i - 1] + (1 - KKSS) * RI[i] * U
        QG[i] = KKG * QG[i - 1] + (1 - KKG) * RG[i] * U

    # 总径流计算
    QT = QI + QG + QS
    return QT


if __name__ == '__main__':
    """
    参考教材： ①包为民《水文预报》（第5版）P144-153
             ②李致家《现代水文模拟与预报技术》（第2版）P35-42
    """
    # --------------第一步，手动更改以下参数，以率定新安江模型，获得最佳参数--------------------------
    # 以下参数需要手动设置
    data = pd.read_excel('代码调试数据.xlsx', sheet_name='Sheet1')
    P = data["模型采用降水（站点2）"].values  # 降雨数据
    Ep = data["模型采用蒸发数据（站点6）"].values  # 蒸发数据
    measured_runoff = data["模型采用径流数据（站点B）"].values  # 洪水数据
    basin_area = 4296.71  # 单位km2
    time_scale = 1  # 小时
    repetitions = 10000  # 率定次数
    calibration_file_name = "xinanjiang_参数率定结果.csv"  # 注意与参数率定中的名称保存一致
    # 参数设置结束

    paras_all_name = ["xinanjiang_W", "xinanjiang_WU", "xinanjiang_WL", "xinanjiang_B", "xinanjiang_C",
                             "xinanjiang_SM", "xinanjiang_EX", "xinanjiang_KG", "xinanjiang_KI", "xinanjiang_KKG",
                             "xinanjiang_KKSS", "xinanjiang_CS", "xinanjiang_FE", "xinanjiang_IMP"
                             ]

    # --------------第二步，绘制预报结果--------------------------
    best_paras_all_values = best_paras_from_name(pd.read_csv(calibration_file_name + ".csv"), param_names=paras_all_name)
    W = best_paras_all_values["xinanjiang_W"]
    WU = best_paras_all_values["xinanjiang_WU"]
    WL = best_paras_all_values["xinanjiang_WL"]
    B = best_paras_all_values["xinanjiang_B"]
    C = best_paras_all_values["xinanjiang_C"]
    SM = best_paras_all_values["xinanjiang_SM"]
    EX = best_paras_all_values["xinanjiang_EX"]
    KG = best_paras_all_values["xinanjiang_KG"]
    KI = best_paras_all_values["xinanjiang_KI"]
    KKG = best_paras_all_values["xinanjiang_KKG"]
    KKSS = best_paras_all_values["xinanjiang_KKSS"]
    CS = best_paras_all_values["xinanjiang_CS"]
    FE = best_paras_all_values["xinanjiang_FE"]
    IMP = best_paras_all_values["xinanjiang_IMP"]
    simulated_runoff = XinAnJiang(WM=W, WUM=WU, WLM=WL, B=B, C=C, SM=SM, EX=EX, KG=KG, KI=KI, KKG=KKG,
                                  KKSS=KKSS, CS=CS, FE=FE, IMP=IMP, P=P, Ep=Ep)
    figure = plt.Figure()
    time = data["时间"].values
    plt.plot(time, simulated_runoff)
    plt.plot(time, measured_runoff)
    plt.legend(["predicted_runoff", "measured_runoff"])
    x_start, x_median, x_final = time[0], time[int(len(time)/2)], time[-1]
    plt.xticks([x_start, x_median, x_final], [str(x_start)[:16], str(x_median)[:16], str(x_final)[:16]])
    plt.title("The predicted runoff", weight="bold")
    plt.show()

    # 将最佳模拟结果输出为excel文件
    data_to_excel = {"时间": time, "模拟径流": simulated_runoff, "实际径流": measured_runoff}
    data_to_excel = pd.DataFrame(data_to_excel)
    data_to_excel.to_excel("xinanjiang_预报最佳结果.xlsx", index=False)
