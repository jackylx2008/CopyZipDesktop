from pathlib import Path


def list_subdirectories(directory):
    for path in Path(directory).rglob("*"):
        if path.is_dir():
            print(path.resolve())


def list_subdirectories_with_keyword(directory, keyword):
    for path in Path(directory).rglob("*"):
        if path.is_dir():
            # 获取路径的最后一个部分
            last_part = path.parts[-1]
            if keyword in last_part:
                print(path.resolve())


# 示例使用
directory_path = (
    "D:/cloudstation/国会二期/12 北京院-主体/415设计变更/415暖通"  # 替换为你的目录路径
)
keyword = "06-03-C2-037-C"  # 替换为你的关键字
# list_subdirectories(base_directory)
list_subdirectories_with_keyword(directory_path, keyword)
