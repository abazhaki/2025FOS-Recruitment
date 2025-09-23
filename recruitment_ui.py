import streamlit as st
import pandas as pd
import os

# 设置页面标题和布局
st.set_page_config(
    page_title="2025FOS-面试信息查询页面",
    layout="wide"
)

# 页面标题
st.title("2025FOS-面试信息查询页面")

def read_encrypted_excel(file_path, password):
    # 创建内存缓冲区
    decrypted = BytesIO()
    
    # 读取加密文件并解密
    with open(file_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        # 验证密码并解密（注意：密码区分大小写）
        office_file.load_key(password=password)
        # 将解密后的内容写入内存缓冲区
        office_file.decrypt(decrypted)
    
    # 将文件指针重置到开头，以便pandas读取
    decrypted.seek(0)
    
    # 用pandas读取解密后的Excel文件
    df = pd.read_excel(decrypted, engine="openpyxl")
    return df
    
# 指定Excel文件路径 - 请在此处修改为你的Excel文件实际路径
EXCEL_FILE_PATH = "results.xlsx"  # 例如: "C:/data/students.xlsx" 或 "./data/results.xlsx"
EXCEL_PASSWORD = st.secrets["DB_PASSWORD"]
st.write("DB password:", st.secrets["db_password"])
# 初始化数据框
df = None

# 尝试读取Excel文件
try:
    # 检查文件是否存在
    if not os.path.exists(EXCEL_FILE_PATH):
        st.error(f"指定路径的文件不存在: {EXCEL_FILE_PATH}")
    else:
        # 读取Excel文件
        # df = pd.read_excel(EXCEL_FILE_PATH)
        df = read_encrypted_excel(EXCEL_FILE_PATH, EXCEL_PASSWORD)
        
        # 检查数据是否包含所需的列
        required_columns = ['姓名', '学号', '结果']
        if not set(required_columns).issubset(df.columns):
            st.error("Excel文件必须包含'姓名'、'学号'和'结果'三列！")
            df = None
        else:
            st.success("感谢您对2025FOS的支持与配合！")
            # 显示数据预览（可选，可注释掉）
            # with st.expander("点击查看数据预览"):
            #     st.dataframe(df)
            
            # 提取所有姓名和学号用于自动完成
            all_names = sorted(df['姓名'].unique())
            all_ids = sorted(df['学号'].astype(str).unique())
            
            # 查询区域
            st.subheader("查询结果")
            col1, col2 = st.columns(2)
            
            with col1:
                # 姓名输入，带自动完成
                name = st.text_input("请输入姓名", placeholder="例如：张三")
            
            with col2:
                # 学号输入，带自动完成
                student_id = st.text_input("请输入学号", placeholder="例如：1001")
            
            # 或者使用文本输入框（如果不需要下拉选择）
            # name = st.text_input("请输入姓名")
            # student_id = st.text_input("请输入学号")
            
            # 查询按钮
            if st.button("查询"):
                if name and student_id:
                    # 转换学号类型以匹配数据框中的类型
                    try:
                        # 尝试将输入的学号转换为与数据框中相同的类型
                        id_type = type(df['学号'].iloc[0])
                        student_id = id_type(student_id)
                    except:
                        pass
                    
                    # 执行查询
                    result = df[(df['姓名'] == name) & (df['学号'] == student_id)]
                    
                    if not result.empty:
                        st.success(f"查询成功！{name}（{student_id}）的结果如下：")
                        st.info(f"结果：{result.iloc[0]['结果']}")
                        st.balloons()
                    else:
                        st.warning("未找到匹配的记录，请检查输入的姓名和学号是否正确")
                else:
                    st.warning("请输入姓名和学号后再查询")
                    
except Exception as e:
    st.error(f"文件处理出错：{str(e)}")
