# 上传 GitHub 并打包分发 · 操作指南

> 目标：代码放 GitHub，**真实数据一个字节都不上传**；打包好的程序放 Releases，换电脑时直接下载解压即用。

---

## 0. 先想清楚两件事

**① 仓库可以公开，数据绝对不行。**

程序本身不含任何个人信息（`data/`、`photos_src/`、`*.db`、成绩/培养方案 xlsx 已在 `.gitignore` 里排除）。但一旦 `git add .` 把 `data/` 提交上去，238 名学生的学号、姓名、谈话记录、成绩就会永久留在 GitHub 历史里，即使之后删掉也能被翻出来——这属于个人信息泄露事故，不只是"不小心"。

所以：

- 公开仓库 = 只放代码 + `data_template/`（空表头）
- 想连数据一起备份 → 用**私有仓库**，且只存代码，数据单独加密备份

**② 安装包不需要"安装"，绿色 zip 更省事。**

本项目是纯 Python 标准库 + 一个静态 HTML，已经用 PyInstaller 打包过 `辅导员工作台.exe`（内置 Python 运行时，双击即用，不用装环境）。所以"安装包"实际上就是：**exe + 源码 + 空数据目录，压成一个 zip**，用户解压双击就行，不往系统里写注册表、不留卸载残留。

---

## 1. 上传代码到 GitHub（3 分钟）

### 1.1 网页建仓库

1. 打开 <https://github.com/new>
2. Repository name：`counselor-workbench`（随便起，避免用真实姓名/学校名）
3. 选 **Private**（推荐；以后想公开再改）或 Public
4. **不要**勾 Add README / .gitignore / license（本地已经有了）
5. Create repository

### 1.2 本地提交并推送

在项目目录打开 Git Bash：

```bash
cd "C:/Users/20690/Downloads/23/23/辅导员工作台_输出"

git init
git add .
git status            # ← 关键：确认列表里没有 data/、photos/、*.db、*.xlsx、*.exe
git commit -m "辅导员工作台 v2.0.0：离线看板 + 学业分析"
git branch -M main
git remote add origin https://github.com/<你的用户名>/counselor-workbench.git
git push -u origin main
```

推送时会要求登录：用 **Personal Access Token** 当密码（GitHub 已不支持账号密码）。

> 生成 token：GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic) → Generate new token，勾 `repo`，复制下来当密码用。

### 1.3 提交前自查（重要）

```bash
git status --short | grep -E "data/|photos|\.db|\.xlsx" && echo "有敏感文件！先检查 .gitignore" || echo "干净"
git ls-files | grep -E "^data/|^photos|\.db$|\.xlsx$" && echo "已入库的敏感文件！" || echo "仓库无敏感文件"
```

如果发现 `data/` 已经被 `git add` 了：

```bash
git rm -r --cached data photos
git commit -m "移除误提交的数据目录"
```

（若已经 push 过，还需要 `git filter-repo` 或直接删库重建，别心存侥幸。）

---

## 2. 打一个"安装包"（1 分钟）

```bash
python make_release.py          # 默认 2.0.0
python make_release.py 2.0.1    # 指定版本号
```

产物：`dist/辅导员工作台-v2.0.0-win64.zip`，内含

```
辅导员工作台-v2.0.0/
├─ 辅导员工作台.exe      双击启动（内置 Python，免安装）
├─ 启动编辑模式.bat      没有 exe 时的备用启动方式
├─ *.py                 全部源码（server / build / academic …）
├─ template.html         前端模板
├─ data/                只有表头的空数据结构
├─ photos_src/  photos/  空目录，放照片用
├─ README.md  使用说明.md  首次使用说明.txt
```

包里**没有**任何学生数据、照片、数据库、导出文件。

---

## 3. 上传到 Releases（让别人能下载）

1. 打开仓库页面 → 右侧 **Releases** → **Create a new release**
2. **Choose a tag**：输入 `v2.0.0` → 点 "Create new tag: v2.0.0 on publish"
3. Release title：`v2.0.0 首个可用版本`
4. 描述里写清：本版本新增学业模块、默认折叠培养方案、支持成绩替代认定等
5. **Attach binaries**：把 `dist/辅导员工作台-v2.0.0-win64.zip` 拖进去
6. Publish release

下载地址形如：

```
https://github.com/<用户名>/counselor-workbench/releases/download/v2.0.0/辅导员工作台-v2.0.0-win64.zip
```

README 里那句 `[Releases](../../releases)` 会自动指向这个页面。

---

## 4. 以后改了功能怎么更新

```bash
python make_release.py 2.0.1     # 重新打包
git add -A
git commit -m "v2.0.1：修复 xxx"
git push
```
然后在 GitHub 上再建一个 `v2.0.1` 的 Release，把新 zip 拖进去。老版本 Release 保留，方便回退。

> 每次改完前端记得走一遍：`build.py` 重建 → `template.html` 同步（template 与 dashboard 去掉 `const DATA` 行后必须完全一致）→ `node --check` 语法校验，再打包。

---

## 5. 换电脑怎么用

1. 新电脑装 Git（可选）或直接下 zip
2. **下 zip**：解压 → 双击 `辅导员工作台.exe` → 空看板起来
3. 数据怎么过去：
   - 首选：把旧电脑的 `data/`、`photos_src/` 整个拷过来覆盖
   - 或者：在新电脑用「导入表格 / 导入成绩 / 导入培养方案」重新导一遍
4. 记得把旧电脑的 `data/` 也加密备份一份（U 盘 / 加密压缩包），别只放 GitHub

---

## 6. 一页速查

| 想做的事 | 做法 |
|---|---|
| 只想给自己换电脑用 | 私有仓库放代码 + 移动硬盘放 data |
| 想分享给同事 | 发 Release 的 zip，不含数据 |
| 想让别人看效果 | 发 `dashboard_脱敏外发版.html`（单文件，双击即开） |
| 想公开做开源项目 | 仓库 Public，但 README 里删掉学校名与真实人名 |
| 担心数据泄露 | 每次 commit 前跑一遍第 1.3 节的自查命令 |
