# 答辩材料生成器

一个专业的法律答辩材料自动生成工具，支持快手和快购两个平台的答辩文书生成。

## ✨ 功能特点

- ✅ 支持快手（北京快手科技有限公司）答辩材料生成
- ✅ 支持快购（成都快购科技有限公司）答辩材料生成
- ✅ 自动生成授权委托书、任职证明、答辩状
- ✅ 订单号可选（留空自动删除）
- ✅ ZIP打包下载，避免浏览器拦截
- ✅ 简洁美观的Web界面

## 📦 安装

### 1. 克隆项目
```bash
git clone https://github.com/wangjingjing08/defense-generator.git
cd defense-generator
```

### 2. 安装依赖
```bash
pip3 install -r requirements.txt
```

### 3. 配置Skills
将快手答辩和快购答辩的skill文件夹复制到 `~/.codeflicker/skills/` 目录：

```
~/.codeflicker/skills/
├── kuaishou-defense/
│   ├── fill_defense.py
│   └── templates/
│       ├── 1-快手-授权委托书.docx
│       ├── 2-快手-任职证明.doc
│       └── 3-快手-答辩状.doc
└── kuaigou-defense/
    ├── fill_defense.py
    └── templates/
        ├── 1-快购-授权委托书.docx
        ├── 2-快购-任职证明.doc
        └── 3-快购-答辩状.doc
```

> **注意**: Skills文件包含模板文档，不包含在仓库中，需要单独配置。

### 4. 启动服务
```bash
python3 server.py
```

访问: http://localhost:8888

## 🚀 使用方式

### 本地使用
```
http://localhost:8888
```

### 局域网共享
```
http://你的IP地址:8888
```

获取IP地址：
```bash
ifconfig | grep "inet " | grep -v 127.0.0.1
```

## 📋 使用说明

1. **选择平台**: 北京快手 或 成都快购
2. **填写案件信息**:
   - 原告（必填）
   - 被告（必填）
   - 案号（必填）
   - 店铺名（必填）
   - 订单编号（选填）
   - 受理法院（必填）
3. **生成下载**: 点击"生成答辩材料"，自动下载ZIP文件

## 🔧 配置

### 修改输出目录
编辑 `server.py`：
```python
OUTPUT_DIR = Path.home() / "Downloads"  # 修改为你的目录
```

### 修改端口
编辑 `server.py`：
```python
app.run(debug=True, host='0.0.0.0', port=8888)  # 修改端口号
```

## 📁 项目结构

```
defense-generator/
├── index.html          # 前端页面
├── server.py           # Flask后端服务
├── requirements.txt    # Python依赖
├── README.md          # 项目文档
├── .gitignore         # Git忽略配置
└── 使用指南.md        # 用户使用指南
```

## ⚠️ 注意事项

1. **Skills依赖**: 必须配置 `kuaishou-defense` 和 `kuaigou-defense` 两个skill
2. **模板文件**: 模板必须包含正确的占位符（【原告】、【被告】等）
3. **权限**: 确保输出目录有写入权限
4. **网络**: 局域网共享需要在同一WiFi网络

## 🐛 常见问题

**Q: 端口被占用？**  
A: 修改 `server.py` 中的端口号，或关闭占用端口的程序

**Q: 生成失败？**  
A: 检查Skills是否正确安装，模板文件是否存在

**Q: 只生成了1个文档？**  
A: 已修复，现在使用ZIP打包下载3个文档

**Q: 局域网无法访问？**  
A: 检查防火墙设置，确保8888端口开放

## 📝 技术栈

- **后端**: Python 3.9+ / Flask 3.1.3
- **前端**: HTML5 / CSS3 / JavaScript (Vanilla)
- **文档处理**: python-docx
- **跨域**: Flask-CORS

## 👥 贡献

欢迎提交Issue和Pull Request！

## 📞 联系方式

- **开发者**: 王晶晶
- **邮箱**: wangjingjing08@kuaishou.com

## 📄 License

MIT License

---

*最后更新: 2026年4月3日*
