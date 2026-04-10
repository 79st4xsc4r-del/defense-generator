#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
答辩材料生成器 - Flask后端服务
"""

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
import subprocess
import os
from pathlib import Path
from datetime import datetime
import json
import zipfile
import tempfile

app = Flask(__name__, static_folder='static', template_folder='.')
CORS(app)

# 配置
OUTPUT_DIR = Path("/tmp")  # 云端使用临时目录

# Skills目录 - 兼容两种结构
SKILLS_DIR_NEW = Path(__file__).parent / "skills"  # 正确结构: skills/kuaishou-defense
SKILLS_DIR_OLD = Path(__file__).parent  # 临时兼容: kuaishou-defense直接在根目录
SKILLS_DIR = SKILLS_DIR_NEW if SKILLS_DIR_NEW.exists() else SKILLS_DIR_OLD

@app.route('/')
def index():
    """首页"""
    return render_template('index.html')

@app.route('/api/generate', methods=['POST'])
def generate_documents():
    """生成答辩文档API"""
    try:
        data = request.json
        
        # 验证必填字段
        required_fields = ['platform', 'plaintiff', 'defendant', 'caseNumber', 'shopName', 'court']
        for field in required_fields:
            if not data.get(field):
                return jsonify({'error': f'缺少必填字段: {field}'}), 400
        
        # 获取参数
        platform = data['platform']
        plaintiff = data['plaintiff']
        defendant = data['defendant']
        case_number = data['caseNumber']
        shop_name = data['shopName']
        order_number = data.get('orderNumber', '')
        court = data['court']
        
        # 确定skill路径
        if platform == 'kuaishou':
            skill_dir = SKILLS_DIR / "kuaishou-defense"
        elif platform == 'kuaigou':
            skill_dir = SKILLS_DIR / "kuaigou-defense"
        else:
            return jsonify({'error': '无效的平台选择'}), 400
        
        if not skill_dir.exists():
            return jsonify({'error': f'Skill不存在: {skill_dir}'}), 500
        
        # 构建命令
        script_path = skill_dir / "fill_defense.py"
        cmd = [
            'python3', str(script_path),
            '--plaintiff', plaintiff,
            '--defendant', defendant,
            '--case-number', case_number,
            '--shop-name', shop_name,
            '--business-entity', defendant,  # 默认使用被告作为经营主体
            '--order-number', order_number,
            '--court', court,
            '--output-dir', str(OUTPUT_DIR)
        ]
        
        # 执行生成
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode != 0:
            return jsonify({
                'error': '生成失败',
                'details': result.stderr
            }), 500
        
        # 查找生成的文件
        today = datetime.now().strftime('%Y%m%d')
        case_safe = case_number.replace('/', '-').replace('(', '').replace(')', '')
        
        generated_files = [
            f"授权委托书_{case_safe}_{today}.docx",
            f"任职证明_{today}.docx",
            f"答辩状_{case_safe}_{today}.docx"
        ]
        
        # 验证文件是否存在
        existing_files = []
        for filename in generated_files:
            file_path = OUTPUT_DIR / filename
            if file_path.exists():
                existing_files.append(filename)
        
        # 创建ZIP文件
        if existing_files:
            zip_filename = f"答辩材料_{case_safe}_{today}.zip"
            zip_path = OUTPUT_DIR / zip_filename
            
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for filename in existing_files:
                    file_path = OUTPUT_DIR / filename
                    zipf.write(file_path, filename)
        
            return jsonify({
                'success': True,
                'message': f'成功生成 {len(existing_files)} 个文档',
                'zip_file': zip_filename,
                'files': existing_files,
                'output_dir': str(OUTPUT_DIR)
            })
        else:
            return jsonify({'error': '未找到生成的文件'}), 500
        
    except subprocess.TimeoutExpired:
        return jsonify({'error': '生成超时'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/download/<filename>')
def download_file(filename):
    """下载文件"""
    try:
        file_path = OUTPUT_DIR / filename
        if not file_path.exists():
            return jsonify({'error': '文件不存在'}), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/health')
def health_check():
    """健康检查"""
    skills = {
        'kuaishou': (SKILLS_DIR / "kuaishou-defense").exists(),
        'kuaigou': (SKILLS_DIR / "kuaigou-defense").exists()
    }
    
    return jsonify({
        'status': 'ok',
        'skills': skills,
        'output_dir': str(OUTPUT_DIR)
    })


if __name__ == '__main__':
    # 从环境变量读取端口（云平台部署）或使用默认8888端口（本地运行）
    port = int(os.environ.get('PORT', 8888))
    
    print("=" * 60)
    print("🚀 答辩材料生成器服务启动")
    print("=" * 60)
    print(f"📁 输出目录: {OUTPUT_DIR}")
    print(f"🔧 Skills目录: {SKILLS_DIR}")
    print(f"🌐 端口: {port}")
    print("=" * 60)
    
    app.run(debug=False, host='0.0.0.0', port=port)
