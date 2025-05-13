from pathlib import Path
import openpyxl
from datetime import datetime
from typing import List, Dict, Any, Optional
import json

'''
定义了 DocumentParser 类，用于解析功能规范 PDF 文件和 CAN 信号矩阵 Excel 文件。
它可以从 PDF 文件中提取结构化的功能需求，
从 Excel 文件中提取 CAN 信号信息，并将这些信息合并返回。
'''

class OutputHandler:
    def __init__(self, config):
        self.config = config
        
    def save_to_excel(self, test_cases: List[Dict[str, Any]], 
                     output_path: Optional[Path] = None) -> Path:
        """保存测试用例到Excel文件"""
        if not test_cases:
            print("❌ 没有测试用例可保存")
            return None
            
        if not output_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = self.config.OUTPUTS_DIR / f"TestCases_{timestamp}.xlsx"
            
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # 表头
        headers = [
            "Object Type", "Name", "Short Description / Action", 
            "Expected Result", "input signal", "output signal",
            "Feature", "Test Group", "Test Case", "Precondition"
        ]
        ws.append(headers)
        
        # # 填充数据
        for case in test_cases:
            # 转换 input_signal 格式
            input_signal_dict = case.get("input_signal", {})
            input_signal_str = ',\n'.join([f"{key}={value}" for key, value in input_signal_dict.items()])
            
            # 转换为标准格式
            function_row = {
                "Object Type": "Function",
                "Name": self._extract_function_name(case.get("description", "")),
                "Short Description / Action": case.get("description", ""),
                "Expected Result": ", ".join(case.get("expected", [])),
                "input signal": input_signal_str,
                "output signal": case.get("output_signal", ""),
                "Feature": self._extract_feature(case.get("coverage", [])),
                "Test Group": self._extract_test_group(case.get("coverage", [])),
                "Test Case": case.get("description", ""),
                "Precondition": "\n".join(case.get("precondition", []))
            }


            
            # 添加功能行
            ws.append([
                function_row["Object Type"],
                function_row["Name"],
                function_row["Short Description / Action"],
                function_row["Expected Result"],
                function_row["input signal"],
                function_row["output signal"],
                function_row["Feature"],
                function_row["Test Group"],
                function_row["Test Case"],
                function_row["Precondition"]
            ])
            
            # 添加测试步骤行
            for step in case.get("steps", []):
                ws.append([
                    "Test Step",
                    function_row["Name"],
                    step,
                    "",
                    "",
                    "",
                    function_row["Feature"],
                    function_row["Test Group"],
                    function_row["Test Case"],
                    ""
                ])
            
            # 添加空行分隔不同测试用例
            ws.append([])
            
        # 保存文件
        wb.save(output_path)
        print(f"✅ 测试用例已保存至: {output_path}")
        return output_path
    
    def _extract_function_name(self, description: str) -> str:
        """从描述中提取功能名称"""
        if "倒车灯" in description:
            return "外灯控制"
        elif "门锁" in description:
            return "门锁控制"
        elif "雨刮" in description:
            return "雨刮控制"
        elif "电源" in description:
            return "电源管理"
        else:
            return "功能控制"
    
    def _extract_feature(self, coverage: List[str]) -> str:
        """从覆盖项中提取特性名称"""
        for item in coverage:
            if "倒车灯" in item:
                return "倒车灯功能"
            elif "门锁" in item:
                return "门锁功能"
            elif "雨刮" in item:
                return "雨刮功能"
            elif "电源" in item:
                return "电源管理功能"
        return "其他功能"
    
    def _extract_test_group(self, coverage: List[str]) -> str:
        """从覆盖项中提取测试组名称"""
        for item in coverage:
            if "倒车灯" in item:
                return "倒车灯"
            elif "门锁" in item:
                return "门锁"
            elif "雨刮" in item:
                return "雨刮"
            elif "电源" in item:
                return "电源"
        return "其他"