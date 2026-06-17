"""
BSBB节点爬虫 GUI 增强版 - 从bsbb.cc网站爬取V2Ray节点信息
优化功能：
- 完善节点解析（端口/SNI/完整URL）
- 增强错误处理和断点续爬
- 节点有效性检测
- 实时统计和进度显示
- 优化Excel格式和报告生成
- 修复编码/时区/路径等问题
- 优化结果报告
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import requests
import re
import base64
import json
import os
import sys
from urllib.parse import unquote, urlparse
from datetime import datetime, timedelta, timezone
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import time

def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

COUNTRY_MAP = {
    'us': 'United States',
    'hk': 'Hong Kong',
    'jp': 'Japan',
    'sg': 'Singapore',
    'kr': 'South Korea',
    'de': 'Germany',
    'fr': 'France',
    'gb': 'United Kingdom',
    'ca': 'Canada',
    'au': 'Australia',
    'nl': 'Netherlands',
    'it': 'Italy',
    'ru': 'Russia',
    'tr': 'Turkey',
    'br': 'Brazil',
    'in': 'India',
    'id': 'Indonesia',
    'my': 'Malaysia',
    'th': 'Thailand',
    'vn': 'Vietnam',
    'ph': 'Philippines',
    'es': 'Spain',
    'se': 'Sweden',
    'ch': 'Switzerland',
    'pl': 'Poland',
    'ua': 'Ukraine',
    'mx': 'Mexico',
    'ae': 'United Arab Emirates'
}

FLAG_TO_COUNTRY = {
    '🇺🇸': 'United States',
    '🇭🇰': 'Hong Kong',
    '🇯🇵': 'Japan',
    '🇸🇬': 'Singapore',
    '🇰🇷': 'South Korea',
    '🇩🇪': 'Germany',
    '🇫🇷': 'France',
    '🇬🇧': 'United Kingdom',
    '🇨🇦': 'Canada',
    '🇦🇺': 'Australia',
    '🇳🇱': 'Netherlands',
    '🇮🇹': 'Italy',
    '🇷🇺': 'Russia',
    '🇹🇷': 'Turkey',
    '🇧🇷': 'Brazil',
    '🇮🇳': 'India',
    '🇮🇩': 'Indonesia',
    '🇲🇾': 'Malaysia',
    '🇹🇭': 'Thailand',
    '🇻🇳': 'Vietnam',
    '🇵🇭': 'Philippines',
    '🇪🇸': 'Spain',
    '🇸🇪': 'Sweden',
    '🇨🇭': 'Switzerland',
    '🇵🇱': 'Poland',
    '🇺🇦': 'Ukraine',
    '🇲🇽': 'Mexico',
    '🇦🇪': 'United Arab Emirates'
}

COUNTRY_NAME_MAP = {
    'United States': '美国',
    'Hong Kong': '香港',
    'Japan': '日本',
    'Singapore': '新加坡',
    'South Korea': '韩国',
    'Germany': '德国',
    'France': '法国',
    'United Kingdom': '英国',
    'Canada': '加拿大',
    'Australia': '澳大利亚',
    'Netherlands': '荷兰',
    'Italy': '意大利',
    'Russia': '俄罗斯',
    'Turkey': '土耳其',
    'Brazil': '巴西',
    'India': '印度',
    'Indonesia': '印尼',
    'Malaysia': '马来西亚',
    'Thailand': '泰国',
    'Vietnam': '越南',
    'Philippines': '菲律宾',
    'Spain': '西班牙',
    'Sweden': '瑞典',
    'Switzerland': '瑞士',
    'Poland': '波兰',
    'Ukraine': '乌克兰',
    'Mexico': '墨西哥',
    'United Arab Emirates': '阿联酋',
    'Taiwan': '台湾',
    'Unknown': '未知'
}

PROTOCOL_DESC = {
    'VLESS': 'VLESS协议',
    'VMESS': 'VMESS协议',
    'TROJAN': 'Trojan协议',
    'SS': 'Shadowsocks',
    'SSR': 'ShadowsocksR',
    'HYSTERIA': 'Hysteria',
    'HYSTERIA2': 'Hysteria2'
}

def get_first_flag_emoji(text):
    if not text:
        return None
    cps = list(text)
    for i in range(len(cps) - 1):
        a = ord(cps[i])
        b = ord(cps[i+1])
        if 0x1F1E6 <= a <= 0x1F1FF and 0x1F1E6 <= b <= 0x1F1FF:
            return cps[i] + cps[i+1]
    return None

def extract_country_from_name(name):
    if not name:
        return None
    cn_country_map = {
        '香港': 'Hong Kong', '日本': 'Japan', '韩国': 'South Korea',
        '法国': 'France', '美国': 'United States', '新加坡': 'Singapore',
        '台湾': 'Taiwan', '英国': 'United Kingdom', '德国': 'Germany',
        '加拿大': 'Canada', '澳大利亚': 'Australia', '俄罗斯': 'Russia',
        '土耳其': 'Turkey', '巴西': 'Brazil', '印度': 'India',
        '印尼': 'Indonesia', '马来西亚': 'Malaysia', '泰国': 'Thailand',
        '越南': 'Vietnam', '菲律宾': 'Philippines', '西班牙': 'Spain',
        '瑞典': 'Sweden', '瑞士': 'Switzerland', '波兰': 'Poland',
        '乌克兰': 'Ukraine', '墨西哥': 'Mexico', '阿联酋': 'United Arab Emirates'
    }
    for cn, en in cn_country_map.items():
        if cn in name:
            return en
    return None

def get_protocol(line):
    if not line:
        return None
    line = line.lower()
    if 'vless://' in line: return 'vless'
    if 'vmess://' in line: return 'vmess'
    if 'trojan://' in line: return 'trojan'
    if 'ss://' in line: return 'ss'
    if 'ssr://' in line: return 'ssr'
    if 'hysteria2://' in line: return 'hysteria2'
    if 'hysteria://' in line: return 'hysteria'
    return None

def get_url(line, proto):
    if not line or not proto:
        return None
    token = proto + '://'
    idx = line.lower().find(token)
    if idx < 0:
        return None
    url = line[idx:].split()[0]
    return url.strip()

def parse_host_port_sni(url, proto):
    res = {'host':'N/A','port':'N/A','sni':'N/A'}
    try:
        if proto == 'vmess':
            b64 = url[8:]
            b64 = b64.replace('_','/').replace('-','+')
            b64 += '=' * ((4-len(b64)%4)%4)
            dec = base64.b64decode(b64).decode()
            d = json.loads(dec)
            res['host'] = d.get('add','N/A')
            res['port'] = str(d.get('port','N/A'))
            res['sni'] = d.get('sni', d.get('host','N/A'))
        else:
            pu = urlparse(url)
            nl = pu.netloc
            if '@' in nl:
                nl = nl.split('@')[-1]
            if ':' in nl:
                h,p = nl.split(':',1)
                res['host']=h
                if p.isdigit():
                    res['port']=p
            else:
                res['host']=nl
            q = {}
            if pu.query:
                for kv in pu.query.split('&'):
                    if '=' in kv:
                        k,v = kv.split('=',1)
                        q[k.lower()]=unquote(v)
            for k in ['sni','servername','host']:
                if k in q:
                    res['sni']=q[k]
                    break
    except:
        pass
    return res

def get_latency_ms(line):
    if not line:
        return None
    m = re.search(r'(\d+)ms', line)
    if m:
        return int(m.group(1))
    return None

def check_node_validity(url):
    if not url:
        return False, "URL无效"
    try:
        p = get_protocol(url)
        info = parse_host_port_sni(url,p)
        if info['host']=='N/A' or info['port']=='N/A':
            return False, "无法解析地址"
        return True, "有效"
    except:
        return False, "检测失败"

class BSBBScrapeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("BSBBScrape GUI v20260522.1148")
        self.root.geometry("950x700")
        self.output_dir = tk.StringVar(value=os.getcwd())
        self.is_running = False
        self.stop_event = threading.Event()
        self.check_validity = tk.BooleanVar(value=True)
        self.keep_excel = tk.BooleanVar(value=False)
        self.create_widgets()

    def create_widgets(self):
        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=1, padx=10,pady=10)

        top = ttk.Frame(main)
        top.pack(fill=tk.X,pady=5)
        ttk.Label(top,text="输出目录：").pack(side=tk.LEFT)
        ttk.Entry(top,textvariable=self.output_dir,width=60).pack(side=tk.LEFT,padx=5)
        ttk.Button(top,text="浏览",command=self.browse).pack(side=tk.LEFT)

        opt = ttk.Frame(main)
        opt.pack(fill=tk.X,pady=2)
        ttk.Checkbutton(opt,text="节点有效性检测",variable=self.check_validity).pack(side=tk.LEFT,padx=5)
        ttk.Checkbutton(opt,text="保留Excel文件",variable=self.keep_excel).pack(side=tk.LEFT,padx=5)

        btns = ttk.Frame(main)
        btns.pack(fill=tk.X,pady=5)
        self.start_btn = ttk.Button(btns,text="开始爬取",command=self.start)
        self.start_btn.pack(side=tk.LEFT,padx=3)
        self.stop_btn = ttk.Button(btns,text="停止",command=self.stop,state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT,padx=3)
        ttk.Button(btns,text="清空日志",command=lambda:self.log_text.delete(1.0,tk.END)).pack(side=tk.LEFT,padx=3)
        ttk.Button(btns,text="打开目录",command=self.open_dir).pack(side=tk.LEFT,padx=3)

        stats = ttk.LabelFrame(main,text="实时统计")
        stats.pack(fill=tk.X,pady=5)
        self.st_lines = ttk.Label(stats,text="处理：0")
        self.st_lines.pack(side=tk.LEFT,padx=10)
        self.st_nodes = ttk.Label(stats,text="节点：0")
        self.st_nodes.pack(side=tk.LEFT,padx=10)
        self.st_valid = ttk.Label(stats,text="有效：0")
        self.st_valid.pack(side=tk.LEFT,padx=10)

        logf = ttk.LabelFrame(main,text="日志")
        logf.pack(fill=tk.BOTH,expand=1,pady=5)
        self.log_text = scrolledtext.ScrolledText(logf,width=100,height=22)
        self.log_text.pack(fill=tk.BOTH,expand=1)
        self.log_text.tag_config('info',foreground='blue')
        self.log_text.tag_config('ok',foreground='green')
        self.log_text.tag_config('err',foreground='red')

        self.prog = ttk.Progressbar(main,mode='determinate')
        self.prog.pack(fill=tk.X,pady=3)

        self.status = ttk.Label(main,text="就绪",relief=tk.SUNKEN,anchor=tk.W)
        self.status.pack(fill=tk.X)

    def log(self,msg,tag='info'):
        def _():
            self.log_text.insert(tk.END,msg+'\n',tag)
            self.log_text.see(tk.END)
        self.root.after(0,_)

    def update_stat(self,lines,nodes,valid):
        def _():
            self.st_lines.config(text=f"处理：{lines}")
            self.st_nodes.config(text=f"节点：{nodes}")
            self.st_valid.config(text=f"有效：{valid}")
        self.root.after(0,_)

    def browse(self):
        d = filedialog.askdirectory()
        if d:
            self.output_dir.set(d)

    def open_dir(self):
        p = self.output_dir.get()
        if os.path.exists(p):
            os.startfile(p)

    def start(self):
        if self.is_running:
            return
        self.is_running = True
        self.stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.status.config(text="运行中")
        threading.Thread(target=self.run,daemon=1).start()

    def stop(self):
        self.stop_event.set()
        self.log("停止中...",'err')

    def run(self):
        try:
            out = self.output_dir.get()
            os.makedirs(out,exist_ok=True)
            self.log("开始下载...")
            urls = ['https://www.bsbb.cc/V2RAY.txt','https://www.bsbb.cc/V22RAY.txt']
            lines = []
            headers={'User-Agent':'Mozilla/5.0'}
            for u in urls:
                try:
                    r = requests.get(u,headers=headers,timeout=15,verify=False)
                    if r.status_code==200:
                        lines.extend(r.text.splitlines())
                        self.log(f"下载成功：{u}")
                except Exception as e:
                    self.log(f"下载失败：{e}",'err')
            lines = [x.strip() for x in lines if x.strip()]
            total = len(lines)
            self.log(f"共 {total} 行")
            nodes = []
            valid_cnt = 0
            for i,l in enumerate(lines):
                if self.stop_event.is_set():
                    break
                self.prog['value'] = (i+1)/total*100
                p = get_protocol(l)
                if not p:
                    continue
                url = get_url(l,p)
                if not url:
                    continue
                country = 'Unknown'
                name = unquote(url.split('#')[-1]) if '#' in url else 'N/A'
                c = extract_country_from_name(name)
                if c:
                    country=c
                else:
                    f = get_first_flag_emoji(l)
                    if f in FLAG_TO_COUNTRY:
                        country=FLAG_TO_COUNTRY[f]
                hp = parse_host_port_sni(url,p)
                lat = get_latency_ms(l)
                valid=True
                msg="有效"
                if self.check_validity.get():
                    valid,msg=check_node_validity(url)
                if valid:
                    valid_cnt+=1
                nodes.append({
                    'Country':country,'Protocol':p.upper(),
                    'Host':hp['host'],'Port':hp['port'],'SNI':hp['sni'],
                    'Name':name,'URL':url,'Latency':f"{lat}ms" if lat else 'N/A',
                    'Valid':valid,'Msg':msg
                })
                self.update_stat(i+1,len(nodes),valid_cnt)
            uniq = []
            seen=set()
            for n in nodes:
                k = n['URL']
                if k not in seen:
                    seen.add(k)
                    uniq.append(n)
            self.log(f"去重后：{len(uniq)} 个",'ok')
            
            # 统计节点相关维度
            protocol_count = {}
            country_count = {}
            for node in uniq:
                # 统计协议分布
                proto = node['Protocol']
                protocol_count[proto] = protocol_count.get(proto, 0) + 1
                
                # 统计国家/地区分布
                country = node['Country']
                country_count[country] = country_count.get(country, 0) + 1
            
            # 计算有效率
            valid_rate = (valid_cnt / len(uniq) * 100) if len(uniq) > 0 else 0
            
            # 生成美化的README.md
            readme = os.path.join(out,'README.md')
            with open(readme,'w',encoding='utf-8') as f:
                report_content = f"""# BSBB 节点爬取统计报告
> 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## 📊 核心统计
| 项目 | 数值 |
|------|------|
| 总爬取行数 | {total} |
| 解析出节点数 | {len(nodes)} |
| 去重后节点数 | {len(uniq)} |
| 有效节点数 | {valid_cnt} |
| 节点有效率 | {valid_rate:.2f}% |

## 🛠️ 协议分布
| 协议类型 | 数量 | 占比 |
|----------|------|------|
"""
                # 写入协议分布表格
                for proto, count in protocol_count.items():
                    ratio = (count / len(uniq) * 100) if len(uniq) > 0 else 0
                    proto_cn = PROTOCOL_DESC.get(proto, proto)
                    report_content += f"| {proto_cn} | {count} | {ratio:.2f}% |\n"
                
                report_content += f"""
## 🌍 节点地区分布
| 地区 | 中文名称 | 数量 | 占比 |
|------|----------|------|------|
"""
                # 写入地区分布表格
                for country, count in sorted(country_count.items()):
                    ratio = (count / len(uniq) * 100) if len(uniq) > 0 else 0
                    country_cn = COUNTRY_NAME_MAP.get(country, country)
                    report_content += f"| {country} | {country_cn} | {count} | {ratio:.2f}% |\n"
                
                report_content += f"""
## 📝 备注
- 数据来源：bsbb.cc
- 检测规则：验证节点地址和端口的可解析性
- 去重规则：基于节点URL完全匹配
- 输出文件：
  - `config.txt`: 所有有效节点的URL列表
  - `node.xlsx`: 节点详细信息（仅勾选"保留Excel文件"时生成）
"""
                f.write(report_content)
            
            # 生成Excel文件
            xlsx = os.path.join(out,'node.xlsx')
            wb=Workbook()
            ws=wb.active
            ws.title="节点"
            ws.append(['国家','协议','主机','端口','SNI','名称','延迟','有效','说明','URL'])
            for n in uniq:
                ws.append([
                    COUNTRY_NAME_MAP.get(n['Country'],n['Country']),
                    n['Protocol'],n['Host'],n['Port'],n['SNI'],
                    n['Name'],n['Latency'], '是' if n['Valid'] else '否',n['Msg'],n['URL']
                ])
            wb.save(xlsx)
            
            # 生成config.txt
            config_path = os.path.join(out,'config.txt')
            with open(config_path,'w',encoding='utf-8') as f:
                config_content = '\n'.join([n['URL'] for n in uniq])
                f.write(config_content)
            
            # ====================== 核心修改 ======================
            # 仅输出纯 Base64 字符串到 b64config.json
            try:
                b64_content = base64.b64encode(config_content.encode("utf-8")).decode("utf-8")
                script_folder = os.path.dirname(os.path.abspath(__file__))
                json_path = os.path.join(script_folder, "b64config.json")
                
                with open(json_path, "w", encoding="utf-8") as f:
                    f.write(b64_content)
                    
                self.log(f"✅ 已生成纯Base64配置：b64config.json", "ok")
            except Exception as e:
                self.log(f"❌ 生成b64config.json失败：{e}", "err")
            # ======================================================
            
            # 删除Excel文件（如果未勾选保留）
            if not self.keep_excel.get():
                os.remove(xlsx)
                
            self.log(f"完成！文件保存在：{out}",'ok')
            self.status.config(text="完成")
            messagebox.showinfo("完成",f"成功爬取：{len(uniq)} 个节点")
        except Exception as e:
            self.log(f"错误：{e}",'err')
        finally:
            self.is_running=False
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.prog['value']=0

def main():
    if sys.platform == 'win32':
        try:
            import ctypes
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(),0)
        except:
            pass
    requests.packages.urllib3.disable_warnings()
    root = tk.Tk()
    BSBBScrapeGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
