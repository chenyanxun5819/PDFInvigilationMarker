"""测试 Excel 模板生成和导入功能"""
import sys
from pathlib import Path
from openpyxl import Workbook, load_workbook

# 添加 app 目录到 Python 路径
app_dir = Path(__file__).parent / "app"
sys.path.insert(0, str(app_dir.parent))

from app.template_generator import create_teacher_list_template
from app.excel_importer import import_config_from_excel
from app.config_store import load_existing_payload


def test_template_generation():
    """测试模板生成"""
    print("=" * 60)
    print("测试 1: 模板生成")
    print("=" * 60)
    
    # 使用项目根目录下的测试输出目录
    test_dir = Path(__file__).parent / "test_output"
    test_dir.mkdir(exist_ok=True)
    template_path = test_dir / "test_template.xlsx"
    
    try:
        # 生成模板
        create_teacher_list_template(template_path)
        print(f"✓ 模板生成成功: {template_path}")
        print(f"✓ 文件大小: {template_path.stat().st_size} 字节")
        
        # 验证文件存在
        assert template_path.exists(), "模板文件不存在"
        print("✓ 文件存在性验证通过")
        
        return template_path
    except Exception as e:
        print(f"✗ 模板生成失败: {e}")
        return None


def test_template_import(template_path):
    """测试模板导入"""
    print("\n" + "=" * 60)
    print("测试 2: 模板导入")
    print("=" * 60)
    
    try:
        config = import_config_from_excel(template_path, base_payload=load_existing_payload())
        
        # 检查结构
        assert "teachers" in config, "config 缺少 'teachers' 键"
        assert "seat_blocks" in config, "config 缺少 'seat_blocks' 键"
        assert "metadata" in config, "config 缺少 'metadata' 键"
        
        print(f"✓ 配置导入成功")
        print(f"✓ 老师数量: {len(config['teachers'])}")
        print(f"✓ 座位区块数: {len(config['seat_blocks'])}")
        
        # 显示示例数据
        if config['teachers']:
            print("\n老师示例数据:")
            for i, teacher in enumerate(config['teachers'][:2]):
                print(f"  [{i+1}] {teacher['initial']} - {teacher['name']} (座号: {teacher['seat_no']}, ID: {teacher['teacher_id']}, 组别: {teacher['group']})")
        
        if config['seat_blocks']:
            print("\n座位区块示例数据:")
            for i, block in enumerate(config['seat_blocks'][:3]):
                print(f"  [{i+1}] 座位 {block['seat_no']} -> {block['block']} 区")
        
        print("\n✓ 所有验证通过!")
        return True
        
    except Exception as e:
        print(f"✗ 模板导入失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_old_format_import():
    """测试旧版只有老师名单工作表的格式"""
    print("\n" + "=" * 60)
    print("测试 3: 旧版老师名单格式匯入")
    print("=" * 60)

    test_dir = Path(__file__).parent / "test_output"
    test_dir.mkdir(exist_ok=True)
    old_format_path = test_dir / "old_format_only_teacher.xlsx"

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "老师名单"
    sheet.append(["代号Initial", "教师姓名 Name", "座号", "教职员编号", "组别"])
    sheet.append(["A01", "张老师", "1", "T001", "语文"])
    workbook.save(old_format_path)

    try:
        config = import_config_from_excel(old_format_path, base_payload=load_existing_payload())
        assert len(config["teachers"]) == 1, "旧格式老师数量不正确"
        assert len(config["seat_blocks"]) > 0, "未沿用现有座位区块资料"
        assert config["metadata"].get("seat_blocks_source") == "existing_config", "座位区块来源标记不正确"
        print(f"✓ 旧格式匯入成功: {old_format_path}")
        print(f"✓ 沿用座位区块数: {len(config['seat_blocks'])}")
        return True
    except Exception as e:
        print(f"✗ 旧格式匯入失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_template_headers(template_path):
    """测试模板栏位是否回到旧版格式"""
    print("\n" + "=" * 60)
    print("测试 4: 模板栏位格式")
    print("=" * 60)

    workbook = load_workbook(template_path, data_only=True)
    sheet = workbook["老师名单"]
    headers = [sheet.cell(row=1, column=index).value for index in range(1, 6)]
    expected = ["代号Initial", "教师姓名 Name", "座号", "教职员编号", "组别"]
    if headers != expected:
        print(f"✗ 模板栏位不符: {headers}")
        return False
    print(f"✓ 模板栏位正确: {headers}")
    return True


if __name__ == "__main__":
    print("\n🧪 PDF 监考工具 - Excel 功能测试\n")
    
    # 测试模板生成
    template_path = test_template_generation()
    
    if template_path:
        # 测试模板导入
        success = test_template_import(template_path)
        success = test_old_format_import() and success
        success = test_template_headers(template_path) and success
        
        if success:
            print("\n" + "=" * 60)
            print("✅ 所有测试通过!")
            print("=" * 60)
        else:
            print("\n❌ 某些测试失败")
            sys.exit(1)
    else:
        print("\n❌ 模板生成失败，无法继续测试")
        sys.exit(1)
