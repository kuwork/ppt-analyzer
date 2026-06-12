# Sed 坐标定位 + 多行替换/插入速查文档

> 适用场景：在 Markdown 索引文件中，按页码精确定位并替换/插入识别概要内容。

---

## 一、基础确认

```bash
# 确认 sed 可用（GNU sed 支持 -i 参数）
sed --version

# 测试简单替换
echo "hello world" | sed 's/world/earth/'
# 输出：hello earth
```

---

## 二、单页定位替换（推荐）

### 场景：替换第 N 页的识别概要占位符

**关键技巧：用页标题作为定位锚点，配合 `---` 分隔符限定范围**

```bash
# 替换第 5 页的"待分析后填入"（只处理从"## 第 5 页"到第一个"---"之间的内容）
sed -i '/## 第 5 页/,/^---$/{s/待分析后填入/【第5页识别概要内容】/}' 幻灯片索引.md
```

**原理：**
- `/## 第 5 页/,/^---$/` — 地址范围：从匹配"## 第 5 页"的行开始，到匹配"---"的行结束
- `{s/old/new/}` — 在该范围内执行替换

---

## 三、多行内容插入（使用临时文件）

### 场景：将多行识别概要插入到指定页

**方法：先写入临时文件，再用 sed 读取临时文件插入**

```bash
# 1. 准备要插入的多行内容（保存到临时文件）
cat > /tmp/summary_page5.txt << 'EOF'
> **识别概要**
>
> 1. **标题区**（顶部左侧）：1.1 项目背景：形式要求
> 2. **三列卡片**（页面主体）：
>    - 左栏：工作要求...
>    - 中栏：业务要求...
>    - 右栏：管理要求...
> 3. **页面类型**：内容页
EOF

# 2. 使用 sed 在"## 第 5 页"之后插入内容
# 方法 A：在匹配行后读取文件插入
sed -i '/## 第 5 页/r /tmp/summary_page5.txt' 幻灯片索引.md

# 方法 B：删除原占位符，再插入（先删后插）
sed -i '/## 第 5 页/,/^---$/{s/待分析后填入//}' 幻灯片索引.md
sed -i '/## 第 5 页/r /tmp/summary_page5.txt' 幻灯片索引.md
```

---

## 四、整页重建（最可靠）

### 场景：完全重写某一页的内容

```bash
# 1. 准备新页面内容
cat > /tmp/page5_new.txt << 'EOF'
## 第 5 页

![第5页](assets/幻灯片_05.png)

**备注**：（备注内容）

> **识别概要**
>
> 1. **标题区**（顶部左侧）：1.1 项目背景：形式要求
> 2. **正文**（标题下方）：...

---
EOF

# 2. 删除旧页面内容（从"## 第 5 页"到下一个"## 第"或文件末尾）
sed -i '/## 第 5 页/,/^---$/{/## 第 5 页/!d; s/.*/PLACEHOLDER/}' 幻灯片索引.md

# 3. 替换占位符为新内容
sed -i '/PLACEHOLDER/r /tmp/page5_new.txt' 幻灯片索引.md
sed -i '/PLACEHOLDER/d' 幻灯片索引.md
```

---

## 五、批量处理脚本模板

```bash
#!/bin/bash
# batch_insert.sh
# 用法：将多页识别概要批量插入索引文件

INDEX_FILE="幻灯片索引.md"
SUMMARY_DIR="./summaries"  # 每页一个文件：page_01.md, page_02.md, ...

for i in $(seq -w 1 41); do
    page_num=$(echo $i | sed 's/^0//')  # 去前导零
    summary_file="$SUMMARY_DIR/page_$i.md"
    
    if [ -f "$summary_file" ]; then
        # 删除该页原有的"待分析后填入"占位符
        sed -i "/## 第 $page_num 页/,/^---$/{s/待分析后填入//}" "$INDEX_FILE"
        
        # 在页标题后插入识别概要
        sed -i "/## 第 $page_num 页/r $summary_file" "$INDEX_FILE"
        
        echo "已插入第 $page_num 页"
    fi
done
```

---

## 六、Python 退化方案（当 sed 不可用时）

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re

def insert_summary(index_path, page_num, summary_text):
    """在索引文件中插入指定页的识别概要"""
    with open(index_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 定位该页范围
    pattern = rf'(## 第 {page_num} 页\n\n!\[第{page_num}页\]\(assets/幻灯片_{page_num:02d}\.png\)\n\n\*\*备注\*\*：.*?\n)(\n> \*\*识别概要\*\*：.*?\n)?(?=\n---|\Z)'
    
    replacement = rf'\1\n{summary_text}\n'
    content = re.sub(pattern, replacement, content, flags=re.DOTALL)
    
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(content)

# 批量处理示例
def rebuild_index(index_path, notes_dict, summaries_dict, total_pages=41):
    """按页码顺序重建完整索引文件"""
    lines = ['# 幻灯片索引', '', '## 目录速览', ...]
    
    for i in range(1, total_pages + 1):
        note = notes_dict.get(i, '')
        summary = summaries_dict.get(i, '> **识别概要**：待补充')
        
        lines.extend([
            f'## 第 {i} 页',
            '',
            f'![第{i}页](assets/幻灯片_{i:02d}.png)',
            '',
            f'**备注**：{note if note else "（无）"}',
            '',
            summary,
            '',
            '---',
            '',
        ])
    
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
```

---

## 七、常见问题

| 问题 | 原因 | 解决 |
|---|---|---|
| sed 替换影响了其他页 | 地址范围没有正确限定 | 确保使用 `/## 第 N 页/,/^---$/` 限定范围 |
| 多行内容插入后格式错乱 | 直接嵌入换行符导致 | 使用临时文件 + `r` 命令读取 |
| 特殊字符导致 sed 报错 | 未转义 `/` 或 `&` | 使用其他分隔符如 `#`：`s#old#new#` |
| 中文内容显示乱码 | 文件编码不一致 | 确保文件为 UTF-8 编码 |

---

## 八、最佳实践总结

1. **优先使用页码标记定位**：每页内容以 `## 第 N 页` 开头，以 `---` 结尾，形成明确的地址范围
2. **单页替换用 sed 地址范围**：`/## 第 N 页/,/^---$/`
3. **多行插入用临时文件**：`sed '/锚点/r 临时文件'`
4. **批量处理用循环脚本**：配合 `seq` 生成页码序列
5. **sed 不可用时退化到 Python**：在内存中构建完整内容后一次性写入
