import re
import pandas as pd

def extract_elements(markdown_content):
    lines = markdown_content.split('\n')
    elements = []

    for line in lines:
        if re.match(r'^#+\s', line):
            title_level = line.count('#')
            title_text = line.strip('#').strip()
            elements.append({'type': 'title', 'level': title_level, 'text': title_text})
        elif re.match(r'^-?\s*定义\s*$', line):
            elements.append({'type': 'definition', 'text': '定义'})
        elif re.match(r'^-?\s*[A-Za-z]+\s*[:：=]\s*[A-Za-z]+', line):
            elements.append({'type': 'proposition', 'text': line})
        else:
            # 计算前导空格数量，用于确定层级
            indent_count = len(re.match(r'^\s*', line).group())
            elements.append({'type': 'other', 'text': line, 'indent': indent_count})

    # 删除所有 "other" 类型且 "text" 为空的元素
    elements = [element for element in elements if element['type'] != 'other' or element['text'].strip()]

    return elements

def analyze_structure(elements):
    structure_analysis = {}

    current_level = 1
    max_title_level = 1  # 用于存储最大的 'title' 类型元素的层级
    parent_stack = []  # 用栈来维护 parent

    for i in range(len(elements)):
        current_element = elements[i]

        if current_element['type'] == 'title':
            # 更新最大的 'title' 类型元素的层级
            max_title_level = max(max_title_level, current_element['level'])
            
            # 判断标题层级
            if 'level' in current_element and i > 0 and current_element['level'] > elements[i - 1].get('level', 0):
                current_level += 1
                # 将当前标题的索引压入栈
                parent_stack.append(i - 1)
            elif 'level' in current_element and i > 0 and current_element['level'] < elements[i - 1].get('level', 0):
                current_level -= 1
                # 弹出栈顶，找到当前标题的 parent
                while parent_stack and ('indent' not in elements[parent_stack[-1]] or elements[parent_stack[-1]]['indent'] >= current_element['level']):
                    parent_stack.pop()

        # 处理 'definition' 类型元素与 'title' 类型元素同级的情况
        elif current_element['type'] == 'definition':
            parent_index = parent_stack[-1] if parent_stack else None
            structure_analysis[i] = {'type': current_element['type'], 'text': current_element['text'], 'level': current_level, 'parent': parent_index}
            continue

        # 根据前导空格数量确定层级
        elif current_element['type'] == 'other' and 'indent' in current_element:
            indent_level = max_title_level + current_element['indent'] + 4  # 用最大 'title' 的层级作为基准
            if indent_level > current_level:
                current_level = indent_level
                # 将当前元素的索引压入栈
                parent_stack.append(i - 1)
            elif indent_level < current_level:
                current_level = indent_level
                # 弹出栈顶，找到当前元素的 parent
                while parent_stack and ('indent' not in elements[parent_stack[-1]] or elements[parent_stack[-1]]['indent'] >= current_element['indent']):
                    parent_stack.pop()

        # parent 是栈顶元素
        parent_index = parent_stack[-1] if parent_stack else None
        structure_analysis[i] = {'type': current_element['type'], 'text': current_element['text'], 'level': current_level, 'parent': parent_index}

    return structure_analysis

# 读取Markdown文件内容
with open(r"C:\Users\23038\Desktop\Converter\形式系统.md", 'r', encoding='utf-8') as file:
    markdown_content = file.read()

# 运行提取函数和结构分析函数
elements = extract_elements(markdown_content)
structure_analysis = analyze_structure(elements)

# 提取parent-children关系
parent_children_data = {'关系': [], '源': [], '目标': []}

for i, analysis in structure_analysis.items():
    parent_text = analysis['text'] if analysis['parent'] is not None else None
    children_texts = [re.sub(r'^\s*-', '', child_analysis['text']).strip() for child_i, child_analysis in structure_analysis.items() if child_analysis['parent'] == i and child_analysis['text'].strip()]
    
    # 如果存在空指向（'parent' 为空或 'children' 为空），则忽略这些关系
    if parent_text is not None and children_texts:
        for child_text in children_texts:
            parent_children_data['关系'].append('继承')  # 使用空字符串填充关系列
            parent_children_data['源'].append(re.sub(r'^\s*-', '', parent_text).strip())
            parent_children_data['目标'].append(re.sub(r'^\s*-', '', child_text).strip())

# 创建包含parent-children关系的DataFrame
parent_children_df = pd.DataFrame(parent_children_data)

# 处理Entities中的每个元素，删除前导空格及其后的连字符（-）
entities_df = pd.DataFrame({'实体名称': [re.sub(r'^\s*-', '', analysis['text']).strip() for analysis in structure_analysis.values()], '类别': '主题'})

# 创建Bi-relations表头
bi_relations_df = pd.DataFrame(columns=['orig', 'forward', 'backward'])

# 将DataFrame写入Excel文件的不同sheet
output_path = r'C:\Users\23038\Desktop\Converter\input.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    entities_df.to_excel(writer, sheet_name='Entities', index=False)
    parent_children_df.to_excel(writer, sheet_name='Relations', index=False)
    bi_relations_df.to_excel(writer, sheet_name='bi-relations', index=False)

print(f"Excel文件已成功保存至: {output_path}")
