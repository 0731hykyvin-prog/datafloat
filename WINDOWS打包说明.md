# Datafloat Windows EXE 打包说明

## 最终效果

打包完成后得到 **单个 exe 文件**：
```
dist\Datafloat数据处理平台.exe
```
将该 exe 拷贝到任意 Windows 10/11 电脑，双击即可运行，**无需安装 Python 或任何依赖**，完全离线可用。

---

## 打包流程（两步）

### 第一步：在联网电脑上下载依赖包

双击运行：
```
_download_dependencies.cmd
```

脚本会自动：
1. 检测 Python 环境
2. 下载所有依赖包（.whl 文件）到 `wheels\` 目录
3. 下载 PyInstaller

> **注意**：下载用的电脑和最终打包电脑，建议使用相同 Python 大版本和相同 64 位系统。
> 例如用 Python 3.11 下载的 cp311 包，就应在 Python 3.11 环境里离线安装和打包。

### 第二步：在内网（离线）电脑上打包

将**整个项目目录**拷贝到内网 Windows 电脑，确保包含：
```
wheels\              ← 第一步下载的离线依赖包
requirements.txt
_build_offline_exe.cmd
Datafloat.spec
main.py
core\
ui\
templates\
```

双击运行：
```
_build_offline_exe.cmd
```

脚本会：
1. 自动检测 Python（优先使用 `portable_python\` 便携版，其次使用系统安装的 Python）
2. 从 `wheels\` 离线安装所有依赖
3. 运行 PyInstaller 打包成独立 exe

---

## 如果没有 Python

如果内网电脑完全没有 Python，可以放入便携版：

```
Datafloat_codex
├─ portable_python
│  ├─ python.exe
│  ├─ python311.dll 等文件
│  └─ Lib
├─ wheels
├─ _build_offline_exe.cmd
└─ ...
```

> Python 大版本必须与 wheels 文件夹中的包匹配：
> - wheels 中有 `cp311` → 用 Python 3.11
> - wheels 中有 `cp312` → 用 Python 3.12

---

## 运行说明

- 程序是纯桌面应用，不需要浏览器和服务器
- `templates` 模板已内置进 exe，首次运行自动复制到可写目录
- 用户在"字段映射"里保存的新模板会保存到 exe 同级目录或 Windows 用户 AppData 目录
- 合并结果默认输出到 exe 同级目录下的 `merged.xlsx`

---

## 常见问题

### Q: 双击 exe 没反应？
临时将 `Datafloat.spec` 中 `console=False` 改为 `console=True`，重新打包后运行可看到错误信息。

### Q: 打包后 exe 很大？
正常。因为 Python 解释器 + PySide6 + pandas + matplotlib 等大型库会被打包进去。约 200-400 MB。

### Q: 能否减小体积？
可以在 `Datafloat.spec` 的 `excludes` 列表中添加不需要的模块，或使用 `upx=True` 压缩（需先安装 UPX）。

### Q: 打包时报 "No module named xxx"？
在 `Datafloat.spec` 的 `hiddenimports` 列表中添加缺失的模块名，重新打包。
