import pandas as pd
import streamlit as st
import openpyxl
from io import BytesIO  # 恢复io模块用于Excel生成

def process_data(data_1, data_2, data_3):
    """数据处理核心逻辑（与原版本一致）"""
    df1 = data_1.copy()
    df2 = data_2.copy()
    df3 = data_3.copy()
    
    # 处理月计划表
    df1['渠道编号'] = df1['渠道编号'].ffill()
    df1['平台渠道'] = df1['平台渠道'].ffill()
    df1['三级产品名称'] = df1['三级产品名称'].ffill()
    
    # 修正资方名称映射
    df1_replace = {'富邦': '富邦华一', '阳光': '阳光消金'}
    df1['资方'] = df1['资方'].replace(df1_replace)
    
    # 渠道计划统计
    channel_stats = df1.groupby('渠道编号').agg(
        has_in_plan=('月初表内计划', lambda x: (x.notnull() & (x >= 0)).any()),
        has_out_plan=('月初表外计划', lambda x: (x.notnull() & (x >= 0)).any())
    ).reset_index()
    
    # 表内表外计划合并
    in_plan_df = df1[df1['月初表内计划'].notnull() & (df1['月初表内计划'] >= 0)].copy()
    in_plan_df['初次基础计划目标'] = in_plan_df['月初表内计划']
    in_plan_df['最新计划目标'] = in_plan_df['最新表内计划']
    
    out_plan_df = df1[df1['月初表外计划'].notnull() & (df1['月初表外计划'] >= 0)].copy()
    out_plan_df['初次基础计划目标'] = out_plan_df['月初表外计划']
    out_plan_df['最新计划目标'] = out_plan_df['最新表外计划']
    
    result_df = pd.concat([in_plan_df, out_plan_df], ignore_index=True)
    result_df = pd.merge(result_df, channel_stats, on='渠道编号', how='left')
    
    # 设置表外标记
    def set_is_off_balance(row):
        if row['has_in_plan'] and row['has_out_plan']:
            return 'N' if (row['初次基础计划目标'] == row['月初表内计划'] or row['最新计划目标'] == row['最新表内计划']) else 'Y'
        elif row['has_in_plan']:
            return 'N'
        elif row['has_out_plan']:
            return 'SLA'
        else:
            return '未知'
    result_df['是否表外'] = result_df.apply(set_is_off_balance, axis=1)
    
    # 合并维表
    result_df = pd.merge(result_df, df2, left_on='渠道编号', right_on='channel_no', how='left')
    result_df = pd.merge(result_df, df3, left_on='资方', right_on='bank_name_map', how='left')
    
    # 字段处理
    result_df.loc[result_df['是否表外'] == 'N', ['bank_id', 'bank_name']] = ['', '']
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
    
    # 填充缺失值和格式调整
    final_df[['初次资金成本', '最新资金成本']] = final_df[['初次资金成本', '最新资金成本']].fillna(0.036)
    final_df[['初次基础计划目标', '最新计划目标']] = final_df[['初次基础计划目标', '最新计划目标']].astype(int)
    
    # **关键修复：确保所有字段正确添加**
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
    
    # 调整列顺序（确保所有字段存在）
    final_df = final_df[[
        '业务类型', '目标类型', '时间', '渠道编号', '渠道名称', '三级产品编码', '三级产品名称', 
        '资方编码', '资方名称', '是否表外', '二级分类', '三级分类', '初次基础计划目标', 
        '最新计划目标', '初次进阶版计划', '表外计划放款上限', '表外计划放款下限', 
        '初次资金成本', '最新资金成本', '初次资产价格', '最新资产价格', '出资比例', 
        '单日放款偏离度', '放款开始时间', '放款结束时间', '1次推送期望占比'
    ]]
    
    return final_df

def to_excel(df):
    """恢复Excel生成函数，确保字段格式正确"""
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title('EXCEL 小工具-月计划')
    
    try:
        dim_file = pd.ExcelFile('./渠道维表.xlsx')
        data_2 = dim_file.parse('渠道')
        data_3 = dim_file.parse('资方')
        # st.success('渠道维表读取成功！')
        
        # 上传月计划数据
        st.header('上传月计划')
        data_1_file = st.file_uploader('月计划数据', type=['xlsx', 'xls'])
        
        if data_1_file:
            data_1 = pd.read_excel(data_1_file, sheet_name='月计划')
            st.success('读取成功')
            
            if st.button('处理'):
                with st.spinner('处理中...'):
                    final_df = process_data(data_1, data_2, data_3)
                
                st.success('完成！')
                st.dataframe(final_df)
                
                # 下载结果（使用恢复的to_excel函数）
                st.header('下载')
                excel_file = to_excel(final_df)
                st.download_button(
                    label='下载Excel文件',
                    data=excel_file,
                    file_name='补录结果.xlsx',
                    mime='补录.sheet'
                )
        else:
            st.info('请上传月计划数据以继续')
            
    except FileNotFoundError:
        st.error('未找到渠道维表文件，请确保"渠道维表.xlsx"在仓库同目录下')
    except KeyError as ke:
        st.error(f'字段缺失错误: {str(ke)}，请检查数据列名是否匹配')
    except Exception as e:
        st.error(f'处理数据时出错: {str(e)}')

if __name__ == '__main__':
    main()
