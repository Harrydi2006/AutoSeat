#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试数据生成器
用于生成座位分配系统的测试数据，包含各种复杂情况
"""

import pandas as pd
import random
import string
from typing import List, Tuple, Dict, Optional
import argparse
import os
from datetime import datetime

class TestDataGenerator:
    """测试数据生成器类"""
    
    def __init__(self, seed: Optional[int] = None):
        """初始化生成器
        
        Args:
            seed: 随机种子，用于生成可重复的测试数据
        """
        if seed is not None:
            random.seed(seed)
        
        # 常用中文姓名库
        self.surnames = [
            '王', '李', '张', '刘', '陈', '杨', '赵', '黄', '周', '吴',
            '徐', '孙', '胡', '朱', '高', '林', '何', '郭', '马', '罗',
            '梁', '宋', '郑', '谢', '韩', '唐', '冯', '于', '董', '萧'
        ]
        
        self.given_names = [
            '伟', '芳', '娜', '秀英', '敏', '静', '丽', '强', '磊', '军',
            '洋', '勇', '艳', '杰', '娟', '涛', '明', '超', '秀兰', '霞',
            '平', '刚', '桂英', '华', '建华', '建国', '建军', '志强', '志明', '秀珍',
            '晓明', '晓红', '小红', '小明', '小华', '小丽', '小燕', '小芳', '小娟', '小静'
        ]
    
    def generate_names(self, count: int) -> List[str]:
        """生成指定数量的随机姓名
        
        Args:
            count: 需要生成的姓名数量
            
        Returns:
            List[str]: 生成的姓名列表
        """
        names = set()
        while len(names) < count:
            surname = random.choice(self.surnames)
            given_name = random.choice(self.given_names)
            name = f"{surname}{given_name}"
            names.add(name)
        
        return list(names)
    
    def generate_student_list(self, count: int, output_file: str) -> List[str]:
        """生成学生名单Excel文件
        
        Args:
            count: 学生数量
            output_file: 输出文件路径
            
        Returns:
            List[str]: 生成的学生姓名列表
        """
        names = self.generate_names(count)
        
        # 创建DataFrame
        df = pd.DataFrame({
            '姓名': names
        })
        
        # 保存到Excel
        df.to_excel(output_file, index=False)
        print(f"学生名单已保存到: {output_file}")
        
        return names
    
    def generate_preferences_data(
        self, 
        names: List[str], 
        willing_levels: int = 3,
        unwilling_levels: int = 3,
        fill_rate_range: Tuple[float, float] = (0.3, 0.8),
        output_file: str = "preferences.xlsx"
    ) -> Dict:
        """生成偏好关系数据
        
        Args:
            names: 学生姓名列表
            willing_levels: 喜好权重等级数量
            unwilling_levels: 不喜好权重等级数量
            fill_rate_range: 填充率范围 (最小值, 最大值)
            output_file: 输出文件路径
            
        Returns:
            Dict: 生成的偏好关系统计信息
        """
        # 创建工作簿
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            
            # 生成喜好关系数据
            willing_data = {}
            willing_stats = {}
            
            for level in range(1, willing_levels + 1):
                col1_name = f"喜好{level}_人员1"
                col2_name = f"喜好{level}_人员2"
                
                # 随机决定这个等级的填充率
                fill_rate = random.uniform(*fill_rate_range)
                max_pairs = min(len(names) // 2, int(len(names) * fill_rate))
                
                # 生成随机的人员对
                pairs = []
                used_names = set()
                
                for _ in range(max_pairs):
                    # 随机选择两个不同的人
                    available_names = [n for n in names if n not in used_names]
                    if len(available_names) < 2:
                        break
                    
                    person1 = random.choice(available_names)
                    available_names.remove(person1)
                    person2 = random.choice(available_names)
                    
                    pairs.append((person1, person2))
                    
                    # 根据随机概率决定是否将这些人标记为已使用
                    if random.random() < 0.7:  # 70%的概率避免重复使用
                        used_names.add(person1)
                        used_names.add(person2)
                
                # 模拟部分填写情况：随机删除一些条目
                if random.random() < 0.3:  # 30%的概率出现部分填写
                    remove_count = random.randint(1, max(1, len(pairs) // 3))
                    pairs = pairs[:-remove_count]
                
                willing_data[col1_name] = [pair[0] for pair in pairs]
                willing_data[col2_name] = [pair[1] for pair in pairs]
                willing_stats[f"level_{level}"] = len(pairs)
            
            # 生成不喜好关系数据
            unwilling_data = {}
            unwilling_stats = {}
            
            for level in range(1, unwilling_levels + 1):
                col1_name = f"不喜好{level}_人员1"
                col2_name = f"不喜好{level}_人员2"
                
                # 不喜好关系通常比喜好关系少
                fill_rate = random.uniform(0.1, 0.4)
                max_pairs = min(len(names) // 3, int(len(names) * fill_rate))
                
                pairs = []
                for _ in range(max_pairs):
                    person1 = random.choice(names)
                    person2 = random.choice([n for n in names if n != person1])
                    pairs.append((person1, person2))
                
                # 模拟部分填写情况
                if random.random() < 0.4:  # 40%的概率出现部分填写
                    remove_count = random.randint(1, max(1, len(pairs) // 2))
                    pairs = pairs[:-remove_count]
                
                unwilling_data[col1_name] = [pair[0] for pair in pairs]
                unwilling_data[col2_name] = [pair[1] for pair in pairs]
                unwilling_stats[f"level_{level}"] = len(pairs)
            
            # 合并所有数据到一个DataFrame
            max_rows = max(
                max([len(v) for v in willing_data.values()] + [0]),
                max([len(v) for v in unwilling_data.values()] + [0])
            )
            
            # 填充数据到相同长度
            all_data = {}
            
            # 添加喜好数据（从A列开始）- 每列包含用逗号分隔的人名对
            col_index = 0
            for level in range(1, willing_levels + 1):
                col1_name = f"喜好{level}_人员1"
                col2_name = f"喜好{level}_人员2"
                
                data1 = willing_data.get(col1_name, [])
                data2 = willing_data.get(col2_name, [])
                
                # 合并为逗号分隔的人名对
                combined_data = []
                for i in range(len(data1)):
                    if data1[i] and data2[i]:
                        combined_data.append(f"{data1[i]},{data2[i]}")
                    else:
                        combined_data.append('')
                
                # 填充到max_rows长度
                combined_data.extend([''] * (max_rows - len(combined_data)))
                
                all_data[chr(ord('A') + col_index)] = combined_data
                col_index += 1
            
            # 添加空列分隔
            all_data[chr(ord('A') + col_index)] = [''] * max_rows
            col_index += 1
            
            # 添加不喜好数据 - 每列包含用逗号分隔的人名对
            for level in range(1, unwilling_levels + 1):
                col1_name = f"不喜好{level}_人员1"
                col2_name = f"不喜好{level}_人员2"
                
                data1 = unwilling_data.get(col1_name, [])
                data2 = unwilling_data.get(col2_name, [])
                
                # 合并为逗号分隔的人名对
                combined_data = []
                for i in range(len(data1)):
                    if data1[i] and data2[i]:
                        combined_data.append(f"{data1[i]},{data2[i]}")
                    else:
                        combined_data.append('')
                
                # 填充到max_rows长度
                combined_data.extend([''] * (max_rows - len(combined_data)))
                
                all_data[chr(ord('A') + col_index)] = combined_data
                col_index += 1
            
            # 创建DataFrame并保存
            df = pd.DataFrame(all_data)
            df.to_excel(writer, sheet_name='偏好关系', index=False)
            
            print(f"偏好关系数据已保存到: {output_file}")
            
            # 返回统计信息
            return {
                'willing_stats': willing_stats,
                'unwilling_stats': unwilling_stats,
                'willing_levels': willing_levels,
                'unwilling_levels': unwilling_levels,
                'total_students': len(names)
            }
    
    def generate_test_suite(
        self,
        output_dir: str = "test_data",
        scenarios: Optional[List[Dict]] = None
    ):
        """生成完整的测试数据套件
        
        Args:
            output_dir: 输出目录
            scenarios: 测试场景配置列表
        """
        if scenarios is None:
            scenarios = [
                {
                    'name': '小班级_标准情况',
                    'student_count': 20,
                    'willing_levels': 2,
                    'unwilling_levels': 2,
                    'fill_rate_range': (0.6, 0.8)
                },
                {
                    'name': '中班级_复杂情况',
                    'student_count': 35,
                    'willing_levels': 4,
                    'unwilling_levels': 3,
                    'fill_rate_range': (0.3, 0.7)
                },
                {
                    'name': '大班级_稀疏关系',
                    'student_count': 50,
                    'willing_levels': 3,
                    'unwilling_levels': 2,
                    'fill_rate_range': (0.2, 0.4)
                },
                {
                    'name': '极端情况_多等级',
                    'student_count': 25,
                    'willing_levels': 6,
                    'unwilling_levels': 5,
                    'fill_rate_range': (0.1, 0.5)
                },
                {
                    'name': '单等级_高密度',
                    'student_count': 15,
                    'willing_levels': 1,
                    'unwilling_levels': 1,
                    'fill_rate_range': (0.8, 0.9)
                }
            ]
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 生成测试报告
        report = []
        report.append(f"测试数据生成报告 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append("=" * 60)
        
        for i, scenario in enumerate(scenarios, 1):
            print(f"\n生成测试场景 {i}: {scenario['name']}")
            
            # 创建场景目录
            scenario_dir = os.path.join(output_dir, f"{i:02d}_{scenario['name']}")
            os.makedirs(scenario_dir, exist_ok=True)
            
            # 生成学生名单
            student_file = os.path.join(scenario_dir, "students.xlsx")
            names = self.generate_student_list(scenario['student_count'], student_file)
            
            # 生成偏好关系
            preferences_file = os.path.join(scenario_dir, "preferences.xlsx")
            stats = self.generate_preferences_data(
                names=names,
                willing_levels=scenario['willing_levels'],
                unwilling_levels=scenario['unwilling_levels'],
                fill_rate_range=scenario['fill_rate_range'],
                output_file=preferences_file
            )
            
            # 生成配置说明
            config_file = os.path.join(scenario_dir, "config.txt")
            with open(config_file, 'w', encoding='utf-8') as f:
                f.write(f"测试场景: {scenario['name']}\n")
                f.write(f"学生数量: {scenario['student_count']}\n")
                f.write(f"喜好等级数: {scenario['willing_levels']}\n")
                f.write(f"不喜好等级数: {scenario['unwilling_levels']}\n")
                f.write(f"填充率范围: {scenario['fill_rate_range']}\n\n")
                
                f.write("建议的单元格范围配置:\n")
                willing_range = f"A1:{chr(ord('A') + scenario['willing_levels'] * 2 - 1)}50"
                unwilling_start = chr(ord('A') + scenario['willing_levels'] * 2 + 1)
                unwilling_end = chr(ord(unwilling_start) + scenario['unwilling_levels'] * 2 - 1)
                unwilling_range = f"{unwilling_start}1:{unwilling_end}50"
                
                f.write(f"喜好关系范围: {willing_range}\n")
                f.write(f"不喜好关系范围: {unwilling_range}\n\n")
                
                f.write("生成的数据统计:\n")
                for level, count in stats['willing_stats'].items():
                    f.write(f"喜好等级{level.split('_')[1]}: {count}对\n")
                for level, count in stats['unwilling_stats'].items():
                    f.write(f"不喜好等级{level.split('_')[1]}: {count}对\n")
            
            # 添加到报告
            report.append(f"\n场景 {i}: {scenario['name']}")
            report.append(f"  学生数量: {scenario['student_count']}")
            report.append(f"  喜好等级: {scenario['willing_levels']}, 不喜好等级: {scenario['unwilling_levels']}")
            report.append(f"  输出目录: {scenario_dir}")
            
            print(f"场景 {i} 生成完成: {scenario_dir}")
        
        # 保存总报告
        report_file = os.path.join(output_dir, "generation_report.txt")
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report))
        
        print(f"\n所有测试数据生成完成！")
        print(f"输出目录: {output_dir}")
        print(f"生成报告: {report_file}")

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='座位分配系统测试数据生成器')
    parser.add_argument('--output-dir', '-o', default='test_data', 
                       help='输出目录 (默认: test_data)')
    parser.add_argument('--seed', '-s', type=int, 
                       help='随机种子，用于生成可重复的数据')
    parser.add_argument('--student-count', '-n', type=int, default=25,
                       help='学生数量 (默认: 25)')
    parser.add_argument('--willing-levels', '-w', type=int, default=3,
                       help='喜好权重等级数 (默认: 3)')
    parser.add_argument('--unwilling-levels', '-u', type=int, default=3,
                       help='不喜好权重等级数 (默认: 3)')
    parser.add_argument('--generate-suite', action='store_true',
                       help='生成完整的测试套件')
    
    args = parser.parse_args()
    
    # 创建生成器
    generator = TestDataGenerator(seed=args.seed)
    
    if args.generate_suite:
        # 生成完整测试套件
        generator.generate_test_suite(output_dir=args.output_dir)
    else:
        # 生成单个测试场景
        os.makedirs(args.output_dir, exist_ok=True)
        
        # 生成学生名单
        student_file = os.path.join(args.output_dir, "students.xlsx")
        names = generator.generate_student_list(args.student_count, student_file)
        
        # 生成偏好关系
        preferences_file = os.path.join(args.output_dir, "preferences.xlsx")
        stats = generator.generate_preferences_data(
            names=names,
            willing_levels=args.willing_levels,
            unwilling_levels=args.unwilling_levels,
            output_file=preferences_file
        )
        
        print(f"\n测试数据生成完成！")
        print(f"学生名单: {student_file}")
        print(f"偏好关系: {preferences_file}")

if __name__ == "__main__":
    main()