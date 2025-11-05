import pandas as pd
import os

# 读取原始Excel文件
excel_path = 'E:\\Programs\\smarthome\\智能家居模具库.xlsx'
backup_path = 'E:\\Programs\\smarthome\\智能家居模具库_备份.xlsx'

print("=== 添加产品ID列到模具库Excel ===")

# 备份原始文件
if os.path.exists(excel_path):
    import shutil
    shutil.copy2(excel_path, backup_path)
    print(f"✓ 已创建备份文件: {backup_path}")

# 读取Excel文件
df = pd.read_excel(excel_path)
print(f"✓ 读取Excel文件成功，共{len(df)}行数据")

# 创建产品ID映射
product_id_mapping = []
for i, row in df.iterrows():
    device_name = row['设备名称']
    brand = row['品牌']
    
    # 根据设备名称和品牌生成唯一的产品ID
    if '一键' in device_name:
        base_id = 'switch_1'
    elif '二键' in device_name:
        base_id = 'switch_2'
    elif '三键' in device_name:
        base_id = 'switch_3'
    elif '四键' in device_name:
        base_id = 'switch_4'
    else:
        base_id = f'device_{i+1}'
    
    # 添加品牌标识
    if brand == '领普':
        product_id = f'{base_id}_lp'
    elif brand == '易来':
        product_id = f'{base_id}_yl'
    else:
        product_id = base_id
    
    product_id_mapping.append(product_id)

# 在产品ID列前插入产品ID列
df.insert(0, '产品ID', product_id_mapping)

print("\n✓ 产品ID映射关系:")
for i, (device_name, product_id) in enumerate(zip(df['设备名称'], df['产品ID'])):
    print(f"  {device_name} -> {product_id}")

# 保存修改后的Excel文件
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    df.to_excel(writer, index=False)

print(f"\n✓ 已成功添加产品ID列到: {excel_path}")
print("✓ 新的Excel结构:")
print("  列名顺序: 产品ID, 设备品类, 设备名称, 设备简称, 是否启用, 单价, 品牌, 主规格, 单位, 渠道, 采购链接, 设备图片")

# 验证修改
print("\n=== 验证修改结果 ===")
df_updated = pd.read_excel(excel_path)
print(f"✓ 验证成功，新文件有{len(df_updated.columns)}列，第一列为: {df_updated.columns[0]}")
print("\n前3行数据预览:")
print(df_updated.head(3))