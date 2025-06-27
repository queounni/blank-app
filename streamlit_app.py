import pandas as pd
import streamlit as st
import warnings
from io import BytesIO

warnings.filterwarnings('ignore')

def process_data(data_1, data_2, data_3):
    df1 = data_1.copy()
    df2 = data_2.copy()
    df3 = data_3.copy()
    
    # 第一步：处理第一张表（月计划）
    # 前三个字段填补缺失值，使用该列上一个非空值
    df1['渠道编号'] = df1['渠道编号'].ffill()
    df1['平台渠道'] = df1['平台渠道'].ffill()
    df1['三级产品名称'] = df1['三级产品名称'].ffill()
    
    # 修正资方名称映射错误
    df1_replace = {
        '富邦': '富邦华一',
        '阳光': '阳光消金'
    }
    df1['资方'] = df1['资方'].replace(df1_replace)
    
    # 为每个渠道判断是否有表内和表外计划
    channel_stats = df1.groupby('渠道编号').agg(
        has_in_plan=('月初表内计划', lambda x: (x.notnull() & (x >= 0)).any()),
        has_out_plan=('月初表外计划', lambda x: (x.notnull() & (x >= 0)).any())
    ).reset_index()

    # 创建表内计划表
    in_plan_df = df1[df1['月初表内计划'].notnull() & (df1['月初表内计划'] >= 0)].copy()
    in_plan_df['初次基础计划目标'] = in_plan_df['月初表内计划']
    in_plan_df['最新计划目标'] = in_plan_df['最新表内计划']

    # 创建表外计划表
    out_plan_df = df1[df1['月初表外计划'].notnull() & (df1['月初表外计划'] >= 0)].copy()
    out_plan_df['初次基础计划目标'] = out_plan_df['月初表外计划']
    out_plan_df['最新计划目标'] = out_plan_df['最新表外计划']

    # 合并表内和表外计划
    result_df = pd.concat([in_plan_df, out_plan_df], ignore_index=True)

    # 根据渠道统计信息添加"是否表外"标记
    result_df = pd.merge(result_df, channel_stats, on='渠道编号', how='left')

    def set_is_off_balance(row):
        if row['has_in_plan'] and row['has_out_plan']:
            # 既有表内也有表外计划
            if row['初次基础计划目标'] == row['月初表内计划'] or row['最新计划目标'] == row['最新表内计划']:
                return 'N'  # 表内计划
            else:
                return 'Y'  # 表外计划
        elif row['has_in_plan'] and not row['has_out_plan']:
            # 只有表内计划
            return 'N'
        elif not row['has_in_plan'] and row['has_out_plan']:
            # 只有表外计划
            return 'SLA'
        else:
            return '未知'  # 理论上不会出现这种情况

    result_df['是否表外'] = result_df.apply(set_is_off_balance, axis=1)

    result_df = pd.merge(result_df, df2, left_on='渠道编号', right_on='channel_no', how='left')
    result_df = pd.merge(result_df, df3, left_on='资方', right_on='bank_name_map', how='left')

    result_df.loc[result_df['是否表外'] == 'N', 'bank_id'] = ''
    result_df.loc[result_df['是否表外'] == 'N', 'bank_name'] = ''

    # 选择并重新命名需要的字段
    final_df = result_df[[
        '渠道编号', 'channel_desc', 'third_prod_cde', 'third_prod_name', 
        'bank_id', 'bank_name', '是否表外', '二级分类', '三级分类',
        '初次基础计划目标', '月初资产价格', '月初资金成本', '最新计划目标', '最新资产价格', '最新资金成本'
    ]]

    final_df.columns = [
        '渠道编号', '渠道名称', '三级产品编码', '三级产品名称', 
        '资方编码', '资方名称', '是否表外', '二级分类', '三级分类',
        '初次基础计划目标', '初次资产价格', '初次资金成本', '最新计划目标', '最新资产价格', '最新资金成本'
    ]

    # 处理资金成本缺失值，填补为3.6%
    final_df['初次资金成本'] = final_df['初次资金成本'].fillna(0.036)
    final_df['最新资金成本'] = final_df['最新资金成本'].fillna(0.036)
    # 计划为整数
    final_df['初次基础计划目标'] = final_df['初次基础计划目标'].astype(int)
    final_df['最新计划目标'] = final_df['最新计划目标'].astype(int)
    # 添加新字段，值为空
    final_df['业务类型'] = '机构'
    final_df['目标类型'] = '月'
    final_df['时间'] = '2025-06'
    final_df['初次进阶版计划'] = ''
    final_df['表外计划放款上限'] = ''
    final_df['表外计划放款下限'] = ''
    final_df['出资比例'] = ''
    final_df['单日放款偏离度'] = ''
    final_df['放款开始时间'] = ''
    final_df['放款结束时间'] = ''
    final_df['1次推送期望占比'] = ''
    # 调整列顺序
    final_df = final_df[[
        '业务类型', '目标类型', '时间', '渠道编号', '渠道名称', '三级产品编码', '三级产品名称', 
        '资方编码', '资方名称', '是否表外', '二级分类', '三级分类', '初次基础计划目标', 
        '最新计划目标', '初次进阶版计划', '表外计划放款上限', '表外计划放款下限', 
        '初次资金成本', '最新资金成本', '初次资产价格', '最新资产价格', '出资比例', 
        '单日放款偏离度', '放款开始时间', '放款结束时间', '1次推送期望占比'
    ]]
    
    return final_df

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title('海尔消金数据处理系统')
    
    # 上传文件区域
    st.header('上传数据文件')
    col1, col2, col3 = st.columns(3)
    
    with col1:
        data_1_file = st.file_uploader('月计划数据', type=['xlsx', 'xls'])
    with col2:
        data_2_file = st.file_uploader('渠道维表(渠道)', type=['xlsx', 'xls'])
    with col3:
        data_3_file = st.file_uploader('渠道维表(资方)', type=['xlsx', 'xls'])
    
    if data_1_file and data_2_file and data_3_file:
        # 读取文件
        try:
            data_1 = pd.read_excel(data_1_file, sheet_name='月计划')
            data_2 = pd.read_excel(data_2_file, sheet_name='渠道')
            data_3 = pd.read_excel(data_3_file, sheet_name='资方')
            
            st.success('文件读取成功！')
            
            # 数据预览
            st.header('数据预览')
            
            st.subheader('月计划数据')
            st.dataframe(data_1.head())
            
            st.subheader('渠道维表(渠道)')
            st.dataframe(data_2.head())
            
            st.subheader('渠道维表(资方)')
            st.dataframe(data_3.head())
            
            # 处理数据按钮
            if st.button('开始处理数据'):
                with st.spinner('数据处理中...'):
                    final_df = process_data(data_1, data_2, data_3)
                
                st.success('数据处理完成！')
                
                # 结果预览
                st.header('处理结果预览')
                st.dataframe(final_df)
                
                # 下载结果
                st.header('下载处理结果')
                excel_file = to_excel(final_df)
                st.download_button(
                    label='下载Excel文件',
                    data=excel_file,
                    file_name='补录结果.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
        except Exception as e:
            st.error(f'处理文件时出错: {str(e)}')
    else:
        st.info('请上传所有三个Excel文件以继续')

if __name__ == '__main__':
    main()    
