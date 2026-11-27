# 变更记录

## v1.0.2（未发布）

### 文档
- 操作 SOP 重构为 V1.0.2：新增「这次要什么、给什么」输入/输出清单（模板改为可选、自带品牌模板项前）、三平台安装流程（豆包工作本机安装、豆包开放平台企业专属上架、ChatGPT）、六项初稿检查、两平台差异表和三条边界；出稿指令拆为带模板与不带模板两条，均可一键复制。

### 豆包适配
- 新增 `platforms/doubao/平台执行说明.md`：模板取用路径、脚本命令、字体与预览差异、`present_files` 交付、飞书在线幻灯片只出逐页制作稿的边界。
- 新增 `platforms/doubao/scripts/render_pages.py`：`soffice --headless` + `pdftoppm` 逐页渲染，支持 `--pages` 局部重渲。
- `platforms/doubao` 之外的共用脚本：`dump_template.py`（模板页位与形状清单）、`make_supplement_xlsx.py`（客户补充信息表，含行高自适应）。
- `check_ppmt_layout.py`：模板默认取 `assets/`，仓库内运行回落 `templates/`，页位缺失给出可读错误。

### 共用
- `SKILL.md` 增加「平台执行」段，指向包内 `references/平台执行说明.md`；描述补充触发场景与「不要用通用 PPT 流程从零设计版式」。
- 「生成评审稿」段改为「优先用用户本次指定的模板，未指定时用技能自带 `assets/…`」。
- 校验段补充「渲染图逐页看完才算校验完成」。

### 打包
- 新增 `build_skill.py`：同一次提交生成 ChatGPT 与豆包两个 ZIP，模板从 `templates/` 复制到 `assets/`，包内 `VERSION` 记录版本与提交号，固定时间戳可复现。
