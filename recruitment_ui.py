import streamlit as st
import pandas as pd
import os
from io import BytesIO
import msoffcrypto

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
    df_23 = pd.read_excel(decrypted, engine="openpyxl",sheet_name='23级终表')
    df_24 = pd.read_excel(decrypted, engine="openpyxl",sheet_name='24级终表')
    return df_23,df_24
    
# 指定Excel文件路径 - 请在此处修改为你的Excel文件实际路径
EXCEL_FILE_PATH = st.secrets["DB_FILEPATH"]  
EXCEL_PASSWORD = st.secrets["DB_PASSWORD"]

# 初始化数据框
df_23 = None
df_24 = None

# 尝试读取Excel文件
try:
    # 检查文件是否存在
    if not os.path.exists(EXCEL_FILE_PATH):
        st.error(f"指定路径的文件不存在: {EXCEL_FILE_PATH}")
    else:
        # 读取Excel文件
        # df = pd.read_excel(EXCEL_FILE_PATH)
        df_23,df_24 = read_encrypted_excel(EXCEL_FILE_PATH, EXCEL_PASSWORD)
        
        # 检查数据是否包含所需的列
        required_columns = ['姓名', '学号', '手机号码','邮箱','通过结果','面试时间','面试地点']
        if not set(required_columns).issubset(df_23.columns):
            st.error("Excel文件23级必须包含'姓名'、'学号'、'手机号码'、'邮箱'、'通过结果'、'面试时间'和'面试地点'七列！")
            df_23 = None
        if not set(required_columns).issubset(df_24.columns):
            st.error("Excel文件24级必须包含'姓名'、'学号'、'手机号码'、'邮箱'、'通过结果'、'面试时间'和'面试地点'七列！")
            df_23 = None
        else:
            st.markdown("""<div style="background-color: #D0ECE7; color: #2c5234; padding: 10px; border-radius: 5px;">
                        感谢你们对2025FOS的关注、支持与配合！<br>
                        请通过本页面查询面试信息（请务必确认输入的信息与FOS申请问卷填写信息保持一致)，具体面试地点和时间
                        如有投递简历但未查询到自己的信息的同学，请及时联系FOS小助手或4位管理层同学，谢谢！""",unsafe_allow_html=True)
            # 显示数据预览（可选，可注释掉）
            # with st.expander("点击查看数据预览"):
            #     st.dataframe(df)
            
            # 提取所有姓名和学号用于自动完成
            # all_names = sorted(df['姓名'].unique())
            # all_ids = sorted(df['学号'].astype(str).unique())
            st.subheader("请选择查询的类别")
            option = st.segmented_control('类别', ['2023级小朋友', '2024级宣传组'], key="option_seg")
            if option == '2023级小朋友':
                df = df_23
            else:
                df = df_24
            # 将df所有列保存为str格式
            for col in df.columns:
                df[col] = df[col].astype(str)
            # 查询区域
            st.subheader("查询结果")
            col1, col2 = st.columns(2)
            with col1:
                # 姓名输入
                name = st.text_input("请输入姓名", placeholder="例如：张三")
            
            with col2:
                # 学号输入
                student_id = st.text_input("请输入学号", placeholder="例如：1001")

            col3, col4 = st.columns(2)
            with col3:
                # 手机号码输入
                tel = st.text_input("请输入手机号码", placeholder="例如：13800138000")
            with col4:
                # 邮箱输入
                email = st.text_input("请输入邮箱", placeholder="例如：abc123@163.com")

            # 或者使用文本输入框（如果不需要下拉选择）
            # name = st.text_input("请输入姓名")
            # student_id = st.text_input("请输入学号")
            
            # 查询按钮
            if st.button("查询"):
                if name and student_id and tel and email:
                    # 转换学号类型以匹配数据框中的类型
                    # try:
                    #     # 尝试将输入的学号转换为与数据框中相同的类型
                    #     id_type = type(df['学号'].iloc[0])
                    #     student_id = id_type(student_id)
                    # except:
                    #     pass
                    
                    # 执行查询
                    result = df[df['姓名'] == name][df['学号'] == student_id][df['手机号码'] == tel][df['邮箱'] == email]     
                    
                    if result.empty == False:
                        st.success(f"查询成功！{name}（{student_id}）的结果如下：")
                        # st.info(f"{result.iloc[0]['通过结果']}")
                        if result.iloc[0]['通过结果'] == '是':
                            st.balloons()
                            st.success("恭喜你通过初试，进入面试环节！")
                            st.info(f"面试时间：{result.iloc[0]['面试时间']}")
                            st.info(f"面试教室：{result.iloc[0]['面试地点']}")
                            st.info(f"{result.iloc[0]['结果']}")
                            if option == '2023级小朋友':
                                st.image(
                                st.secrets["23WECHAT_QRCODE_BASE64"],  # 从Secrets读取Base64
                                caption="请扫码加入微信群，群内将同步面试注意事项",
                                use_container_width=False,
                                width=150
                                )
                                st.warning("⚠️ 仅限本人加入，请勿转发二维码")
                            else:
                                st.image(
                                st.secrets["24WECHAT_QRCODE_BASE64"],  # 从Secrets读取Base64
                                caption="请扫码加入微信群，群内将同步面试注意事项",
                                use_container_width=False,
                                width=150
                                )
                                st.warning("⚠️ 仅限本人加入，请勿转发二维码")
                            
                        if result.iloc[0]['通过结果'] == '否':
                            st.info(f"{result.iloc[0]['结果']}")
                    else:
                        st.warning("未找到匹配的记录，请检查输入的姓名和学号是否正确")
                else:
                    st.warning("请输入姓名和学号后再查询")
                    
except Exception as e:
    st.error(f"文件处理出错：{str(e)}")
