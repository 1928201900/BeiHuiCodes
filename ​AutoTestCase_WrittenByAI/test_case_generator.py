from openai import OpenAI
from typing import List, Dict, Any, Optional
import json

'''
定义了 TestCaseGenerator 类，其主要功能是根据输入的功能需求和 CAN 信号矩阵生成汽车电子测试用例。
具体流程包括构建提示词、调用 AI 生成测试用例、解析 AI 响应等，
当没有检测到功能需求时，会生成示例测试用例。
'''


class TestCaseGenerator:
    def __init__(self, config):
        self.config = config
        self.client = OpenAI(
            api_key=config.API_KEY,
            base_url=config.BASE_URL
        )
        
    def generate(self, requirements: List[Dict[str, Any]], 
                 signals: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """根据需求和信号生成测试用例"""
        if not requirements:
            print("⚠️ 警告：没有检测到功能需求，将使用示例测试用例")
            return self._generate_example_test_cases(signals)
            
        # 构建提示词
        prompt = self._build_prompt(requirements, signals)
        
        # 调用AI生成测试用例
        response = self._call_ai(prompt)
        if not response:
            return []
            
        # 解析AI响应
        try:
            test_cases = self._parse_response(response, requirements, signals)
            return test_cases
        except Exception as e:
            print(f"⚠️ 解析AI响应失败: {str(e)}")
            print(f"原始响应内容:\n{response[:500]}...")
            return []
    
    def _build_prompt(self, requirements: List[Dict[str, Any]], 
                     signals: Dict[str, Dict[str, Any]]) -> str:
        """构建AI生成测试用例的提示词"""
        # 示例测试用例格式说明
        example_format = """
【测试用例格式】[
  {
    "description": "挂R档，倒车灯点亮",
    "coverage": ["需求1.1", "VCU_ActGear"],
    "input_signal": {
      "IGN1_RELAY_FB": "ON",
      "VCU_ActGear": "0x9",
      "VCU_ActGear_VD": "0x0"
    },
    "output_signal": "倒车灯点亮",
    "precondition": [
      "车辆处于ON电源模式",
      "当前档位为P档",
      "档位信号有效"
    ],
    "steps": [
      "1. 确认车辆处于ON电源模式",
      "2. 确认当前档位为P档",
      "3. 切换档位至R档",
      "4. 观察倒车灯状态"
    ],
    "expected": [
      "倒车灯应点亮"
    ]
  }
]        """
        
        # 构建提示词
        prompt = f"""
你是一位专业的汽车电子测试工程师，擅长根据功能规范和CAN信号矩阵生成全面的测试用例。

【功能规范】
{json.dumps(requirements, indent=2, ensure_ascii=False)}

【CAN信号矩阵】
{json.dumps(signals, indent=2, ensure_ascii=False)}

请基于以上信息，生成符合以下要求的测试用例：
1. 测试用例描述简洁，直接说明测试场景和验证内容，如“挂R档，倒车灯点亮”。
2. 覆盖所有功能需求，包括正常、异常和边界情况
3. 每个测试用例必须包含：
   - 测试描述（简明扼要地说明测试内容）
   - 覆盖的需求ID/信号ID
   - 详细测试步骤（使用序号开头，如"1. 操作内容"）
   - 预期结果（明确的预期行为）
   - 输入信号（使用信号名称和有效值）
   - 输出信号（预期的信号变化）
   - 前置条件（执行测试前必须满足的条件）
4. 输出格式：[
  {{
    "description": "测试描述",
    "coverage": ["需求ID/信号ID"],
    "input_signal": {{
      "信号名称": "信号值"
    }},
    "output_signal": "输出信号变化",
    "precondition": ["前置条件1", "前置条件2"],
    "steps": ["步骤1", "步骤2"],
    "expected": ["预期结果1", "预期结果2"]
  }}
]
以下是一个测试用例的示例格式：
{example_format}
        """
        
        return prompt
    
    def _call_ai(self, prompt: str) -> Optional[str]:
        """调用AI生成测试用例"""
        try:
            print("正在调用AI生成测试用例...")
            completion = self.client.chat.completions.create(
                model=self.config.MODEL,
                messages=[
                    {"role": "system", "content": "你是一个专业的汽车电子测试工程师，擅长根据功能规范和CAN信号矩阵生成全面的测试用例。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,  # 降低随机性，提高确定性
                max_tokens=4000
            )
            return completion.choices[0].message.content
        except Exception as e:
            print(f"⚠️ API调用失败: {str(e)}")
            return None
    
    def _parse_response(self, response: str, 
                       requirements: List[Dict[str, Any]], 
                       signals: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """解析AI响应并转换为内部格式"""
        # 提取JSON部分
        try:
            # 尝试多种JSON解析方式
            start_idx = response.find('[')
            end_idx = response.rfind(']') + 1
            if start_idx < 0 or end_idx <= start_idx:
                # 尝试作为对象解析
                start_idx = response.find('{')
                end_idx = response.rfind('}') + 1
                if start_idx >= 0 and end_idx > start_idx:
                    json_str = response[start_idx:end_idx]
                    data = json.loads(json_str)
                    if isinstance(data, dict) and "cases" in data:
                        return data["cases"]
                    else:
                        return [data] if isinstance(data, dict) else []
                else:
                    raise ValueError("无法找到有效的JSON结构")
            else:
                json_str = response[start_idx:end_idx]
                return json.loads(json_str)
                
        except Exception as e:
            print(f"⚠️ 解析AI响应失败: {str(e)}")
            print(f"原始响应内容:\n{response[:500]}...")
            return []
    
    def _generate_example_test_cases(self, signals: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """生成示例测试用例（当没有检测到需求时使用）"""
        example_cases = [
            {
                "description": "挂R档，倒车灯点亮",
                "coverage": ["示例需求", "VCU_ActGear"],
                "input_signal": {
                    "IGN1_RELAY_FB": "ON",
                    "VCU_ActGear": "0x9",
                    "VCU_ActGear_VD": "0x0"
                },
                "output_signal": "倒车灯点亮",
                "precondition": [
                    "车辆处于ON电源模式",
                    "当前档位为P档",
                    "档位信号有效"
                ],
                "steps": [
                    "1. 确认车辆处于ON电源模式",
                    "2. 确认当前档位为P档",
                    "3. 切换档位至R档",
                    "4. 观察倒车灯状态"
                ],
                "expected": [
                    "倒车灯应点亮"
                ]
            },
            {
                "description": "下OFF电，倒车灯熄灭",
                "coverage": ["示例需求", "IGN1_RELAY_FB"],
                "input_signal": {
                    "IGN1_RELAY_FB": "OFF",
                    "VCU_ActGear": "0x9",
                    "VCU_ActGear_VD": "0x0"
                },
                "output_signal": "倒车灯熄灭",
                "precondition": [
                    "车辆处于ON电源模式",
                    "当前档位为R档",
                    "倒车灯处于点亮状态"
                ],
                "steps": [
                    "1. 确认车辆处于ON电源模式",
                    "2. 确认当前档位为R档且倒车灯点亮",
                    "3. 将车辆切换至OFF电源模式",
                    "4. 观察倒车灯状态"
                ],
                "expected": [
                    "倒车灯应熄灭"
                ]
            },
            {
                "description": "挂N档，倒车灯熄灭",
                "coverage": ["示例需求", "VCU_ActGear"],
                "input_signal": {
                    "IGN1_RELAY_FB": "ON",
                    "VCU_ActGear": "0xA",
                    "VCU_ActGear_VD": "0x0"
                },
                "output_signal": "倒车灯熄灭",
                "precondition": [
                    "车辆处于ON电源模式",
                    "当前档位为R档",
                    "倒车灯处于点亮状态"
                ],
                "steps": [
                    "1. 确认车辆处于ON电源模式",
                    "2. 确认当前档位为R档且倒车灯点亮",
                    "3. 切换档位至N档",
                    "4. 观察倒车灯状态"
                ],
                "expected": [
                    "倒车灯应熄灭"
                ]
            },
            {
                "description": "挂D档，倒车灯熄灭",
                "coverage": ["示例需求", "VCU_ActGear"],
                "input_signal": {
                    "IGN1_RELAY_FB": "ON",
                    "VCU_ActGear": "0x1",
                    "VCU_ActGear_VD": "0x0"
                },
                "output_signal": "倒车灯熄灭",
                "precondition": [
                    "车辆处于ON电源模式",
                    "当前档位为R档",
                    "倒车灯处于点亮状态"
                ],
                "steps": [
                    "1. 确认车辆处于ON电源模式",
                    "2. 确认当前档位为R档且倒车灯点亮",
                    "3. 切换档位至D档",
                    "4. 观察倒车灯状态"
                ],
                "expected": [
                    "倒车灯应熄灭"
                ]
            },
            {
                "description": "挂P档，倒车灯熄灭",
                "coverage": ["示例需求", "VCU_ActGear"],
                "input_signal": {
                    "IGN1_RELAY_FB": "ON",
                    "VCU_ActGear": "0xB",
                    "VCU_ActGear_VD": "0x0"
                },
                "output_signal": "倒车灯熄灭",
                "precondition": [
                    "车辆处于ON电源模式",
                    "当前档位为R档",
                    "倒车灯处于点亮状态"
                ],
                "steps": [
                    "1. 确认车辆处于ON电源模式",
                    "2. 确认当前档位为R档且倒车灯点亮",
                    "3. 切换档位至P档",
                    "4. 观察倒车灯状态"
                ],
                "expected": [
                    "倒车灯应熄灭"
                ]
            },
            {
                "description": "挡位无效，挂R档，倒车灯不点亮",
                "coverage": ["示例需求", "VCU_ActGear_VD"],
                "input_signal": {
                    "IGN1_RELAY_FB": "ON",
                    "VCU_ActGear": "0x9",
                    "VCU_ActGear_VD": "0x1"
                },
                "output_signal": "倒车灯不点亮",
                "precondition": [
                    "车辆处于ON电源模式",
                    "当前档位为P档",
                    "档位信号无效"
                ],
                "steps": [
                    "1. 确认车辆处于ON电源模式",
                    "2. 确认当前档位为P档",
                    "3. 模拟档位信号无效状态",
                    "4. 切换档位至R档",
                    "5. 观察倒车灯状态"
                ],
                "expected": [
                    "倒车灯不应点亮"
                ]
            }
        ]
        return example_cases