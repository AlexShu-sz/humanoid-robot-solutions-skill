# 人形机器人方案自动化

使用 Excel 现场调研表、PPT 模板和现场证据，生成可编辑的人形机器人方案评审稿。正式下载见 [最新发布](../../releases/latest)。

## 操作步骤

| 步骤 | 输入 | 输出 |
|---|---|---|
| 1. 填写现场调研表 | 空白 Excel、项目资料 | 填好的现场调研 Excel |
| 2. 上传资料并生成方案 | 填好的 Excel、指定 PPT 模板、已有视频或图片 | 可编辑的评审稿 PPT，按需附客户补充信息表 |
| 3. 检查初稿 | 评审稿 PPT、补充信息表 | 核对过的评审稿或修订清单 |
| 4. 回传补充并更新 | 回填表、新资料或修改意见 | 修订后的评审稿及按需更新的补充表 |

在可加载技能并支持生成 PPTX 的平台启用技能，将 Excel、PPT 模板和其他现场资料上传到同一对话，再发送以下指令。

> 请使用人形机器人解决方案技能，根据我上传的现场调研 Excel，按上传的 PPT 模板生成可编辑的人形机器人方案评审稿。结合现场资料核对信息，影响方案判断的待确认事项另附 Excel。交付文件默认用中文命名。

## 文件清单

| 用途 | 文件或位置 |
|---|---|
| PPT 模板 | [人形机器人方案_PPT模板_V1.0.0.pptx](templates/人形机器人方案_PPT模板_V1.0.0.pptx) |
| 空白 Excel | [项目现场调研_Excel模板_V1.0.0.xlsx](templates/项目现场调研_Excel模板_V1.0.0.xlsx) |
| 操作 SOP | [四步操作说明](docs/) |
| 技能 | [技能源文件](skill/SKILL.md)，[发布包从最新发布下载](../../releases/latest) |

PPT 模板保留原有 PPMT 示例内容。出稿时按本次项目资料替换封面、客户背景、现场图和工艺数据，逐页复核固定标题、场景图位、时间轴和标准 BOM。

## 仓库结构

| 位置 | 内容 | 维护约定 |
|---|---|---|
| `skill/` | 唯一的 `SKILL.md`、三份共用规则、通用脚本（`check_ppmt_layout.py`、`dump_template.py`、`make_supplement_xlsx.py`） | 业务规则只改这里 |
| `platforms/doubao/` | 豆包执行说明、依赖本机环境的 `render_pages.py` | 只放执行差异，不复制共用内容 |
| `templates/` | PPT 与 Excel 模板各一份 | 升版时只替换这里，不在别处留副本 |
| `build_skill.py` | 从同一次提交生成两个平台 ZIP | 打包与体积门禁 |

`SKILL.md` 的「平台执行」段引用 `references/平台执行说明.md`；该文件由打包时从 `platforms/<平台>/` 注入，仓库里不重复存放。ChatGPT 侧执行说明待补，缺该文件时技能按共用规则和宿主现有能力执行。

## 构建与发布

```bash
python3 build_skill.py --version 1.0.2                      # dist/ 下生成两个 ZIP
python3 build_skill.py --version 1.0.2 --platform doubao    # 只出豆包版
python3 build_skill.py --platform doubao --no-assets        # 开放平台上传版，不含模板
```

- 两个 ZIP 根目录都是 `industrial-robot-solution/`，`SKILL.md` 位于包根，技能名保持 `industrial-robot-solution`。
- 模板从 `templates/` 复制到包内 `assets/`，带模板的包约 10 MB。
- 包内 `VERSION` 记录版本号、提交号与平台；ZIP 使用固定时间戳，同一提交可复现。
- 豆包开放平台上传限整包 16 MB、单文件 2 MB。PPT 模板单文件约 9.9 MB，超限，上传平台时用 `--no-assets`，模板另行分发；工作台本地安装用带模板的包。
- 同一次提交的两个 ZIP 作为同一个 Release 的两个附件发布，标签用小写 `v主.次.修订`。

## 版本维护

1. 从 `main` 开临时分支（如 `doubao-adapt`），通过 Pull Request 审核后合入；不为平台长期维护分支，避免每次同步都重写提交历史。
2. 共用规则与业务判断只改 `skill/`；平台差异只改 `platforms/<平台>/`；模板只在 `templates/` 升版，保持「业务对象_文件用途_V主.次.修订.扩展名」，改名时同步更新引用。
3. 检查空白 Excel 与 PPT 模板可打开；用脱敏调研表和原模板复跑，核对封面、客户背景、人工与机器人动作对应、场景图、时间轴和标准 BOM。
4. 跑 `python3 build_skill.py`，确认两个 ZIP 可解压、`SKILL.md` 在包根、`assets/` 内模板完整；豆包包至少实测「ZIP 能导入并调用」和「同一份脱敏 Excel 与原模板能导出可编辑 PPTX」。
5. 合并后打标签、发 Release，说明附件版本、变更与验证状态。
6. 历史客户方案、视频、报价和项目专用归档保存在有权访问的内部项目目录。
