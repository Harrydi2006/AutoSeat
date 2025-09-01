#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速测试数据生成脚本
用于快速生成座位分配系统的测试数据
"""

from test_data_generator import TestDataGenerator
import os

def main():
    """主函数 - 生成测试数据"""
    print("座位分配系统 - 测试数据生成器")
    print("=" * 40)
    
    # 创建生成器
    generator = TestDataGenerator(seed=42)  # 使用固定种子确保可重复
    
    # 询问用户选择
    print("\n请选择生成模式:")
    print("1. 快速生成 - 生成完整测试套件（推荐）")
    print("2. 自定义生成 - 自定义参数生成单个测试")
    
    choice = input("\n请输入选择 (1 或 2): ").strip()
    
    if choice == "1":
        # 生成完整测试套件
        print("\n正在生成完整测试套件...")
        generator.generate_test_suite(output_dir="test_data")
        
    elif choice == "2":
        # 自定义生成
        print("\n自定义测试数据生成")
        print("-" * 20)
        
        try:
            student_count = int(input("学生数量 (默认25): ") or "25")
            willing_levels = int(input("喜好权重等级数 (默认3): ") or "3")
            unwilling_levels = int(input("不喜好权重等级数 (默认3): ") or "3")
            
            output_dir = input("输出目录 (默认custom_test): ") or "custom_test"
            
            print(f"\n正在生成测试数据...")
            print(f"学生数量: {student_count}")
            print(f"喜好等级: {willing_levels}, 不喜好等级: {unwilling_levels}")
            
            # 创建输出目录
            os.makedirs(output_dir, exist_ok=True)
            
            # 生成学生名单
            student_file = os.path.join(output_dir, "students.xlsx")
            names = generator.generate_student_list(student_count, student_file)
            
            # 生成偏好关系
            preferences_file = os.path.join(output_dir, "preferences.xlsx")
            stats = generator.generate_preferences_data(
                names=names,
                willing_levels=willing_levels,
                unwilling_levels=unwilling_levels,
                fill_rate_range=(0.3, 0.7),
                output_file=preferences_file
            )
            
            # 生成配置说明
            config_file = os.path.join(output_dir, "config.txt")
            with open(config_file, 'w', encoding='utf-8') as f:
                f.write(f"自定义测试数据配置\n")
                f.write(f"学生数量: {student_count}\n")
                f.write(f"喜好等级数: {willing_levels}\n")
                f.write(f"不喜好等级数: {unwilling_levels}\n\n")
                
                f.write("建议的单元格范围配置:\n")
                willing_range = f"A1:{chr(ord('A') + willing_levels * 2 - 1)}50"
                unwilling_start = chr(ord('A') + willing_levels * 2 + 1)
                unwilling_end = chr(ord(unwilling_start) + unwilling_levels * 2 - 1)
                unwilling_range = f"{unwilling_start}1:{unwilling_end}50"
                
                f.write(f"喜好关系范围: {willing_range}\n")
                f.write(f"不喜好关系范围: {unwilling_range}\n\n")
                
                f.write("生成的数据统计:\n")
                for level, count in stats['willing_stats'].items():
                    f.write(f"喜好等级{level.split('_')[1]}: {count}对\n")
                for level, count in stats['unwilling_stats'].items():
                    f.write(f"不喜好等级{level.split('_')[1]}: {count}对\n")
            
            print(f"\n测试数据生成完成！")
            print(f"输出目录: {output_dir}")
            print(f"学生名单: {student_file}")
            print(f"偏好关系: {preferences_file}")
            print(f"配置说明: {config_file}")
            
        except ValueError:
            print("输入错误，请输入有效的数字！")
            return
        except Exception as e:
            print(f"生成过程中出现错误: {e}")
            return
    
    else:
        print("无效选择，程序退出。")
        return
    
    print("\n" + "=" * 40)
    print("数据生成完成！")
    print("\n使用说明:")
    print("1. 将生成的 students.xlsx 作为学生名单上传")
    print("2. 将生成的 preferences.xlsx 作为偏好关系上传")
    print("3. 根据 config.txt 中的建议配置单元格范围")
    print("4. 调整权重滑块进行座位分配优化")

if __name__ == "__main__":
    main()