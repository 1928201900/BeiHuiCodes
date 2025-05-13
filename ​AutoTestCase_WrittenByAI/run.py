from config import Config
from document_parser import DocumentParser
from test_case_generator import TestCaseGenerator
from output_handler import OutputHandler

'''
项目的入口文件，定义了 main 函数。
程序运行时，会依次完成解析文档、生成测试用例和保存结果的操作，
同时会输出相应的执行信息和错误信息。
'''

def main():
    print("🚀 汽车电子测试用例生成系统 v2.0 (通用框架)")
    config = Config()
    
    try:
        # 1. 解析文档
        print("🔧 正在解析输入文档...")
        parser = DocumentParser(config)
        requirements = parser.parse_pdf()
        signals = parser.parse_excel()
        
        print(f"📑 识别到 {len(requirements)} 条功能需求")
        print(f"📶 识别到 {len(signals)} 个CAN信号")
        
        # 2. 生成测试用例
        print("⚡ 正在智能生成测试用例...")
        generator = TestCaseGenerator(config)
        test_cases = generator.generate(requirements, signals)
        
        if not test_cases:
            print("❌ 未能生成任何测试用例")
            return
            
        print(f"✅ 成功生成 {len(test_cases)} 个测试用例")
        
        # 3. 保存结果
        output_handler = OutputHandler(config)
        output_path = output_handler.save_to_excel(test_cases)
        
        if output_path:
            print(f"\n📊 生成统计:")
            print(f"- 生成测试用例数量: {len(test_cases)}")
            print(f"- 输出文件: {output_path}")
        
    except Exception as e:
        print(f"❌ 执行过程中发生错误: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()