from autogluon.tabular import TabularDataset, TabularPredictor
import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
from sklearn.model_selection import learning_curve

# 加载训练集和测试集
# train_data = pd.read_excel('/home/lhy/prediction/train_data.xlsx')
# test_data = pd.read_excel('/home/lhy/prediction/test_data.xlsx')
# train_data = pd.read_excel('/home/lhy/prediction/train_data_contact.xlsx')
# test_data = pd.read_excel('/home/lhy/prediction/test_data_contact.xlsx')

# train_data = pd.read_excel('/home/lhy/prediction/train_data_cdorg.xlsx')
# test_data = pd.read_excel('/home/lhy/prediction/test_data_cdorg.xlsx')

# train_data = pd.read_excel('/home/lhy/prediction/train_cd.xlsx')
# test_data = pd.read_excel('/home/lhy/prediction/test_cd.xlsx')
# train_data = pd.read_excel('/home/lhy/prediction/(7-3)NewThreshold_TriangulationFeature_train.xlsx')
train_data = pd.read_excel('(labeled)rber_pots_communityfeature.xlsx')
# test_data = pd.read_excel('/home/lhy/prediction/(7-3)NewThreshold_TriangulationFeature_test.xlsx')
# train_data = TabularDataset('/home/lhy/prediction/(病人标签)(60特征)完整NBCD社区发现数据.csv')
# train_data = TabularDataset('/home/lhy/prediction/完整NBCD社区发现数据.csv')
# test_data = TabularDataset('/home/lhy/prediction/(病人标签)完整NBCD社区发现数据9-1测试集.csv')

# train_data = pd.read_excel('/home/lhy/prediction/feature_cd.xlsx')
# train_data = pd.read_excel('/home/lhy/prediction/(labeled)rber_pots_communityfeature.xlsx')
# train_data = pd.read_excel('/home/lhy/prediction/Panel2/没有分社区（细胞接触）/Panel2_nocd(con).xlsx')
# test_data = pd.read_excel('/home/lhy/prediction/test.xlsx')
# train_data = TabularDataset('/home/lhy/prediction/NBCD.csv')
# train_data = TabularDataset('/home/lhy/prediction/(病人标签+删除无标签)(无细胞数和百分比)完整NBCD社区发现数据.csv')

# 创建并训练模型
# predictor = TabularPredictor(label='label').fit(train_data, num_gpus=2, presets='best_quality')

# predictor = TabularPredictor(label='label',eval_metric=['accuracy', 'roc_auc', 'f1', 'precision', 'recall', 'mcc', 'balanced_accuracy']).fit(train_data, num_gpus=2)
# autogluon设置交叉验证划分数据集
# predictor = TabularPredictor(label='label').fit(train_data, num_gpus=2, time_limit=600, presets='best_quality', num_bag_folds=5, num_stack_levels=1)
# 训练模型)
predictor = TabularPredictor(label='label',eval_metric='accuracy').fit(train_data, num_gpus=2, presets='best_quality', num_bag_folds=5, num_stack_levels=1)
# 使用原来训练好的模型
# predictor = TabularPredictor.load('/home/lhy/prediction/AutogluonModels/ag-20240429_134342')

# 在测试集上进行预测
# predictions = predictor.predict(test_data)
predictions = predictor.predict(train_data)

# # 评估模型性能
# performance = predictor.evaluate(test_data)
performance = predictor.evaluate(train_data)
print(performance)

# moper = predictor.leaderboard(test_data)
moper = predictor.leaderboard(train_data)
print(moper)
#保存moper
moper.to_csv('/home/lhy/prediction/moper.csv')

# 查看特征重要性
# feature_importance = predictor.feature_importance(test_data)
# feature_importance = predictor.feature_importance(train_data)
# print(feature_importance)

# results = predictor.fit_summary(show_plot=True)