# yueji-geshizhuanhua 分镜脚本转换器（开源脱敏版）

把甲方原始稿（docx / xlsx / 自由文本）一键转换为标准化的短视频分镜拍摄脚本 xlsx。

## 解决什么问题

影像制作者经常收到两类稿件：

- **甲方稿**：只有"主题 + 风格 + 时长 + 几段故事"，没有镜头概念，颗粒度粗
- **制作稿**：需要固定 10 列结构（镜号 / 拍摄地点 / 画面内容 / 镜头设计 / 台词·旁白 / 时长 / 备注 / 道具 / 音效 / 参考画面），能直接交给摄制组、演员、甲方三方对稿

本 skill 把前者自动转换成后者。

## 核心特性

- **旁白铁律**：原稿旁白一字不改，按语义分段，每段旁白匹配对应画面
- **固定视觉规范**：微软雅黑、#383838 深灰标题行、#D9EAD3 浅绿幕标题、白/浅灰交替镜头行、全表文字居中
- **打印即用**：横向 A3 纸张、适合一页、水平居中，Ctrl+P 直接打印对稿
- **拆镜规则**：识别幕结构、判断镜头切分粒度、补充镜头设计与道具备注

## 安装

```bash
# Claude Code
git clone https://github.com/zds07/yueji-geshizhuanhua.git \
  ~/.claude/skills/yueji-geshizhuanhua

# Codex
git clone https://github.com/zds07/yueji-geshizhuanhua.git \
  ~/.codex/skills/yueji-geshizhuanhua

# 通用 Agents
git clone https://github.com/zds07/yueji-geshizhuanhua.git \
  ~/.agents/skills/yueji-geshizhuanhua
```

## 文件结构

```
SKILL.md                     # skill 主入口：触发条件、制作流程、判断规则
references/
  format-spec.md             # 格式规范（列结构 / 视觉 / 打印设置）
  conversion-rules.md         # 甲方稿→制作稿的 17 条转换规则
  polish-rules.md             # 出稿前自检规则（台词/道具/备注的高频返工点）
  examples.md                 # 输入稿件的变体分类与转换路径对照（脱敏通用版）
scripts/
  build_xlsx.py               # 生成 xlsx 的核心脚本
  parse_brief.py              # 解析甲方稿的辅助脚本
  extract_docx_text.py        # 零依赖读取 docx（zipfile + ElementTree，无需 python-docx）
  insert_shots.py             # 在成品 xlsx 上原地插镜并重排编号（不覆盖手工改动）
assets/
  template_structure.json     # xlsx 模板字段定义
```

## 快速开始

```bash
# 方式一：直接调用生成脚本
python scripts/build_xlsx.py --input 结构化数据.json --output 分镜脚本.xlsx

# 方式二：在支持 skill 的 AI 助手（WorkBuddy / Claude Code）中说
# "把这份原稿做成制作版脚本" 或 "按金标准格式出分镜"
```

## 制作流程（6 步）

1. **预处理原稿**：.docx / .xlsx 直接读；.doc 老格式先用 textutil 转 txt
2. **提取不可动的部分**：视频名称、主题、旁白原文（逐字提取，一字不改）
3. **拆镜**：识别幕结构 → 按旁白语义分段 → 每段旁白匹配一组画面 → 补镜头设计、时长、备注、道具
4. **遵守三条铁律**：
   - 旁白原话不动
   - 没给参考图 → 参考画面列留空
   - 口播/采访标注"受访者自由发挥，我方文案仅供参考"
5. **生成 xlsx**：调用 `build_xlsx.py`，套用固定视觉规范 + 打印设置
6. **交付 → 迭代**：用户说"改哪一镜" → 修改对应数据 → 重新生成

## 使用场景

| 稿件类型 | 处理方式 |
|---|---|
| 银行/政企宣传片脚本 | 转换为标准 10 列分镜 |
| 访谈类脚本（某人物访谈片） | 采访问答按固定机位拆镜 |
| 旅游/文化类脚本（某非遗纪录片） | 幕结构按篇章拆分 |
| 反诈/合规短片 | 跨场景必须拆镜，补合规备注 |

## 许可

个人使用，欢迎参考格式规范与转换规则。

## 关于「开源脱敏版」

本仓库是 **脱敏开源版**：只包含规则、结构、脚本与模板，**不包含任何客户项目数据**。

- ❌ 删掉了：真实客户的脚本文件名、人物角色名、商户 / 门店名、地名
- ✅ 保留了：10 列固定结构、7 行信息头规范、幕结构规则、A-roll / B-roll 拆法、采访提问稿处理、数据与信息点写法，以及 4 个可直接运行的 Python 脚本
- 📌 `references/examples.md` 保留全部**变体判断逻辑**（v1-v5 / 旧格式 / 初稿 / 原始稿），只把文件名换成「示例 A / B / C」

**产出效果一致**：按 `assets/template_structure.json` 构造输入，跑 `scripts/build_xlsx.py`，就能得到横向 A3、绿白配色的制作版 xlsx——与原版同一套代码。

想让 skill 更贴合自己的项目习惯，把你自己的历史脚本放进本机私有副本的 `references/` 目录即可，不必（也不应该）上传到公网仓库。
