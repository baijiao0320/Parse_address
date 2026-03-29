import pandas as pd
import cpca
import os
import time

def process_addresses_from_excel():
    """
    读取“合并地址.xlsx”文件中的地址数据，将其拆分为省市区和剩余详细地址，
    并将结果保存回原Excel文件的右侧。
    """
    excel_file_name = "合并地址.xlsx"
    
    # 检查Excel文件是否存在
    if not os.path.exists(excel_file_name):
        print(f"错误：未在当前目录下找到文件 '{excel_file_name}'。")
        print("请确保 '合并地址.xlsx' 与此脚本在同一文件夹中。")
        input("\n按任意键退出。")
        return

    try:
        # 读取Excel文件
        df_original = pd.read_excel(excel_file_name)

        if df_original.empty:
            print(f"错误：文件 '{excel_file_name}' 为空。请确保文件中包含地址数据。")
            input("\n按任意键退出。")
            return
        
        # 尝试识别地址列
        address_column_name = None
        if df_original.shape[1] == 1:
            address_column_name = df_original.columns[0]
            print(f"检测到单列 '{address_column_name}' 作为地址列。")
        else:
            # 优先查找常见的地址列名
            possible_address_columns = ['地址', '详细地址', 'Address', 'full_address']
            for col in possible_address_columns:
                if col in df_original.columns:
                    address_column_name = col
                    break
            
            if address_column_name:
                print(f"检测到多列。已自动选择列 '{address_column_name}' 作为地址列。")
            else:
                # 如果没有常见列名，则默认使用第一列
                address_column_name = df_original.columns[0]
                print(f"检测到多列，但未找到常见地址列名。默认使用第一列 '{address_column_name}' 作为地址列。")
                print("如果地址列不正确，请手动修改脚本中的 'address_column_name' 变量。")

        # 确保地址列的数据类型为字符串，并处理缺失值
        addresses_series = df_original[address_column_name].astype(str).fillna('')

        # 检查是否有实际的地址数据可以处理
        if addresses_series.empty or all(addresses_series == ''):
            print(f"警告：地址列 '{address_column_name}' 中没有有效的地址数据可供处理。")
            input("\n按任意键退出。")
            return

        print("开始解析地址数据，这可能需要一些时间...")
        
        # 使用 cpca 库解析地址
        parsed_addresses_df = cpca.transform(addresses_series.tolist())
        
        # 将 cpca 返回的 '地址' 列重命名为 '剩余详细地址'
        parsed_addresses_df.rename(columns={'地址': '剩余详细地址'}, inplace=True)

        print("地址解析完成。")

        # 将解析后的数据与原始数据合并
        df_final = pd.concat([df_original, parsed_addresses_df], axis=1)

        # 尝试保存文件，处理文件被占用的情况
        max_retries = 5
        for i in range(max_retries):
            try:
                df_final.to_excel(excel_file_name, index=False, engine='openpyxl')
                print(f"\n成功处理地址并保存结果到 '{excel_file_name}'。")
                print("您现在可以打开该文件查看新增的省、市、区和剩余详细地址列。")
                break # 成功保存，退出重试循环
            except PermissionError:
                if i < max_retries - 1:
                    print(f"警告：文件 '{excel_file_name}' 正在被其他程序占用，无法写入。")
                    print(f"请确保该文件已关闭。将在 5 秒后重试... (第 {i+1} 次尝试)")
                    time.sleep(5)
                else:
                    print(f"\n错误：经过多次尝试，文件 '{excel_file_name}' 仍然无法写入。")
                    print("请务必关闭该Excel文件，然后再次运行脚本。")
            except Exception as save_e:
                print(f"\n保存文件时发生未知错误：{save_e}")
                break # 发生其他错误，不再重试

    except FileNotFoundError:
        print(f"错误：文件 '{excel_file_name}' 未找到。请确保它在脚本的同一目录下。")
    except KeyError as e:
        print(f"错误：Excel文件中未找到地址列 '{e}'。请检查列名是否正确或修改脚本中的列名识别逻辑。")
    except pd.errors.EmptyDataError:
        print(f"错误：文件 '{excel_file_name}' 为空。请确保文件中包含数据。")
    except Exception as e:
        print(f"发生了一个未知错误：{e}")
        print("请确保 'pandas'、'cpca' 和 'openpyxl' 库已正确安装 (`pip install pandas cpca openpyxl`)。")
        print("同时，检查 '合并地址.xlsx' 是否是一个有效的Excel文件。")
    
    input("\n按任意键退出。") # 保持控制台窗口打开，以便用户查看输出

if __name__ == "__main__":
    print("地址处理脚本正在启动...")
    process_addresses_from_excel()
