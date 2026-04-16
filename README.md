# 排班管理 - 云打包说明

## 自动化构建

本项目配置了 GitHub Actions，可自动构建 Windows .exe 和 macOS 应用。

### 构建产物

| 平台 | 文件 | 说明 |
|------|------|------|
| Windows | `dist/排班管理/` | 双击 `排班管理.exe` 运行 |
| macOS | `dist/排班管理/排班管理` | 双击运行 |

### 手动部署步骤（推荐）

**第一次设置（仅需1次）：**

1. **创建 GitHub 仓库**
   - 访问 https://github.com/new
   - Repository name 填: `scheduler`
   - 选择 Private（私有）
   - 不要勾选任何初始化选项
   - 点击 Create repository

2. **推送代码到 GitHub**（在终端执行）：
   ```bash
   cd ~/学习爱好/编程/排班管理
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/你的用户名/scheduler.git
   git push -u origin main
   ```

3. **触发构建**
   - 访问: `https://github.com/你的用户名/scheduler/actions`
   - 点击 "Build Windows .exe" 工作流
   - 点击 "Run workflow" → 运行

4. **下载构建产物**
   - 构建完成后点击 job → Artifacts
   - 下载 `排班管理-Windows` 或 `排班管理-macOS`

5. **发布 Release（可选）**
   - 打一个 tag 触发自动构建：
     ```bash
     git tag v1.0
     git push origin v1.0
     ```
   - 或在 GitHub Releases 页面手动创建

### 以后更新代码

```bash
git add .
git commit -m "描述你的修改"
git push
```

GitHub Actions 会自动重新构建。

### 文件说明

- `.github/workflows/build.yml` - 自动化构建配置
- `scheduler.spec` - PyInstaller 打包配置
- `scheduler.py` - 主程序
- `scheduler.db` - 数据库文件

### 本地测试

```bash
# macOS 本地运行
open dist/排班管理/排班管理

# Windows 本地打包（需要 Windows 系统）
pip install pyinstaller pyqt5 openpyxl
pyinstaller scheduler.spec --clean
```
