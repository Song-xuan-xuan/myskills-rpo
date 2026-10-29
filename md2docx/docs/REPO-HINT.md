# 放在多 skill 仓库里怎么组织

本 skill 目录自身是自包含的，推荐仓库结构：

```
your-skills-repo/
├── README.md                 # 仓库总说明：这个仓库放哪些 skill
├── LICENSE
├── md2docx/                  # ← 本目录原样放进来
│   ├── SKILL.md              # pi 要求：<skill 名>/SKILL.md
│   ├── ...
└── other-skill/
    └── SKILL.md
```

安装方式二选一：

```bash
# ① 全局：所有项目可用
git clone <repo> ~/.pi/agent/skills-repo
ln -s ~/.pi/agent/skills-repo/md2docx ~/.pi/agent/skills/md2docx

# ② 项目内：只有该项目可用
mkdir -p .agents/skills
ln -s <repo>/md2docx .agents/skills/md2docx
```

装完在新会话里说 `/skill:md2docx --list-profiles` 验证。

## 提交前检查

```bash
cd md2docx
python3 selftest.py          # 136 项，应全绿
git status --short           # 不应出现 *.docx / *.pdf / build/ / __pycache__
```
