# myskills

我自己的 pi skills 合集。

每个子目录是一个独立、可单独安装的 skill，解决一件具体的事。不是代码库，也不是框架——就是一组写好复用的小工具。

## 什么是 skill

pi（以及兼容的 agent 框架）会从固定目录加载 skill。一个 skill 由两部分组成：

| 文件 | 给谁看 | 内容 |
|---|---|---|
| `SKILL.md` | **agent** | 何时触发、怎么做、有哪些必须遵守的规矩 |
| `README.md` | **人** | 完整手册：原理、用法、坑、限制 |

`SKILL.md` 前面有一段 `description`，agent 靠它判断「当前任务该不该用这个 skill」。所以 `SKILL.md` 写得短而锐，`README.md` 写得全而慢。

## 包含的 skill

| skill | 一句话 |
|---|---|
| [`md2docx`](md2docx/) | 用 Markdown 写报告，一键出排版规范、编号可自动重排的 Word(docx)：图表公式自动题注与交叉引用、封面、分节页码、参考文献、格式档案（技术报告 / 公文 / 学位论文） |
| [`service-monitor`](service-monitor/) | 跑长任务时让手机替你盯着：任意「需要等结果」的任务完成后推一条通知到手机，脚本自带心跳、超时、异常终止检测，跑完自己退出 |

## 安装

```bash
git clone https://github.com/Song-xuan-xuan/myskills-rpo.git ~/myskills
```

然后把想要的那几个 skill **软链接**到 pi 的加载目录。用软链接而不是复制，这样 `git pull` 之后立刻生效。

```bash
# 方式一：全局安装（所有项目都能用）
mkdir -p ~/.pi/agent/skills
ln -s ~/myskills/md2docx         ~/.pi/agent/skills/md2docx
ln -s ~/myskills/service-monitor ~/.pi/agent/skills/service-monitor

# 方式二：只在某个项目里用
cd <你的项目>
mkdir -p .agents/skills
ln -s ~/myskills/md2docx         .agents/skills/md2docx
ln -s ~/myskills/service-monitor .agents/skills/service-monitor
```

装完检查一下（新会话里）：

```
/skill:md2docx --list-profiles
```

能看到输出就说明加载成功了。

> 只装你需要的。skill 太多会占满 agent 的上下文预算，反而让每个 skill 都判断不准。

## 更新

```bash
cd ~/myskills && git pull
```

软链接指向仓库，pull 完就是新版。已经开着的会话需要 `/reload` 才会重新读取。

## 依赖

每个 skill 的依赖写在自己的 `README.md` 里，按需安装：

| skill | 需要 |
|---|---|
| `md2docx` | [pandoc](https://pandoc.org/installing.html) + `pip install -r md2docx/requirements.txt` |
| `service-monitor` | bash 4+、`curl`、`tmux` |

## 仓库约定

加新 skill 时遵守这几条：

- 一个目录 = 一个 skill，目录名用小写连字符（`md2docx`、`service-monitor`）
- 目录里必须有 `SKILL.md`（给 agent）和 `README.md`（给人）
- `SKILL.md` 第一段 `description` 要说清楚**什么时候该用**，这是触发的唯一依据
- 脚本放在 `scripts/`，模板放在 `templates/`，示例放 `examples/`
- **不放任何密钥、频道名、内部地址、真实用户数据**——这是公开仓库
- 依赖写进各自的 `requirements.txt` 或 `README.md`，不要指望在这个根 README 里统一交代

## 许可

暂未指定开源许可。如需他人自由使用，建议补一个 `LICENSE`（例如 MIT）。
