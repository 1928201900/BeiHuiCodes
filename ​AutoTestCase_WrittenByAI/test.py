from config import Config
from document_parser import DocumentParser
from pathlib import Path

def inspect_pdf_content():
    """查看PDF解析内容的调试工具"""
    print("🔍 PDF内容检查工具")
    config = Config()
    parser = DocumentParser(config)
    
    try:
        # 解析PDF并获取原始内容
        raw_content = parser.parse_pdf()
        
        # 创建输出文件路径
        output_path = Path(config.OUTPUTS_DIR) / "pdf_content.txt"
        
        # 写入文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("📄 PDF解析内容报告\n\n")
            for i, item in enumerate(raw_content, 1):
                f.write(f"【条目 {i}】\n")
                f.write(f"类型: {item.get('type', '未知')}\n")
                f.write(f"描述: {item.get('description', '无描述')}\n")
                f.write("内容:\n")
                f.write(item.get('content', '无内容') + "\n\n")
                
        print(f"\n✅ 解析内容已保存至: {output_path}")
        
        # 同时在控制台显示部分内容
        print("\n控制台预览:")
        for i, item in enumerate(raw_content[:3], 1):  # 只显示前3条
            print(f"\n【条目 {i}】")
            print(f"类型: {item.get('type', '未知')}")
            print(f"描述: {item.get('description', '无描述')}")
            print("内容预览:")
            print(item.get('content', '无内容')[:200] + "...")
            
    except Exception as e:
        print(f"❌ 解析失败: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    inspect_pdf_content()