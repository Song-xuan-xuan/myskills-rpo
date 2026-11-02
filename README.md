# myskills

> ⚠️ **精简版，测试中。** `md2docx` 于 2026-11-02 做过一次重构（12 个脚本合并为 6 个、
> 命令收成单入口）。自检 137 项通过，但尚未在真实报告上完整验收。
> 稳定版见 [commit 605bc54](https://github.com/Song-xuan-xuan/myskills-rpo/tree/605bc54)。
> 测完删掉本提示。

我自己的 pi skills 合集。每个子目录是一个独立的 skill，可以直接链接到
`~/.pi/agent/skills/` 或项目内的 `.agents/skills/` 使用。

## 包含的 skill

| skill | 用途 |
|---|---|
| [`md2docx`](md2docx/) | 把 Markdown 报告转成排版规范、编号可自动重排的 Word(docx)：图表公式自动题注与交叉引用、封面、分节页码、参考文献、格式档案（技术报告/公文/学位论文） |

## 安装

```bash
git clone https://github.com/Song-xuan-xuan/myskills-rpo.git ~/myskills

# 方式一：全局（所有项目可用）
mkdir -p ~/.pi/agent/skills
ln -s ~/myskills/md2docx ~/.pi/agent/skills/md2docx

# 方式二：只在某个项目里用
cd <你的项目>
mkdir -p .agents/skills
ln -s ~/myskills/md2docx .agents/skills/md2docx
```

装完在新会话里验证：

```
/skill:md2docx --list-profiles
```

## 更新

```bash
cd ~/myskills && git pull
```

软链接指向仓库，pull 之后即生效；已开着的会话需要 `/reload`。

## 依赖

各 skill 的依赖写在自己的 `requirements.txt` / README 里。
`md2docx` 需要 [pandoc](https://pandoc.org/installing.html) 和 `python-docx`：

```bash
brew install pandoc          # macOS
pip install -r md2docx/requirements.txt
```

## 许可

暂未指定开源许可。如需他人自由使用，建议补一个 `LICENSE`（例如 MIT）。
