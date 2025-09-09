#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
老绅控屏眼 - 多屏幕控制软件
功能：控制和监控多个屏幕，特别适用于医院CT室等场景
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from PIL import Image, ImageTk, ImageGrab
import win32api
import win32con
import win32gui
from screeninfo import get_monitors
import keyboard

import winreg
import configparser
import os
import sys

class ToolTip:
    """工具提示类"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)
    
    def on_enter(self, event=None):
        """鼠标进入时显示提示"""
        if self.tooltip:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(self.tooltip, text=self.text, 
                        background="#ffffe0", relief="solid", borderwidth=1,
                        font=("./29华康宋体W3.ttf", 9))
        label.pack()
    
    def on_leave(self, event=None):
        """鼠标离开时隐藏提示"""
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

class ScreenController:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("老绅控屏眼 - 多屏幕控制")
        self.root.geometry("800x450")
        self.root.resizable(True, True)
        
        # 配置文件路径 - 使用程序所在目录确保EXE兼容性
        if getattr(sys, 'frozen', False):
            # 如果是打包的exe文件
            app_dir = os.path.dirname(sys.executable)
        else:
            # 如果是脚本文件
            app_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(app_dir, "screen_config.ini")
        self.config = configparser.ConfigParser()
        
        # 获取屏幕信息
        self.monitors = self.get_screen_info()
        
        # 加载配置
        self.load_config()
        
        # 创建界面
        self.create_widgets()
        
        # 启动预览更新线程
        self.preview_running = True
        self.preview_thread = threading.Thread(target=self.update_previews, daemon=True)
        self.preview_thread.start()
        
        # 启动自动熄屏检测定时器
        self.auto_screen_timer_running = True
        self.start_auto_screen_timer()
        
        # 注册全局快捷键
        self.setup_global_hotkeys()
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            try:
                self.config.read(self.config_file, encoding='utf-8')
                # 加载检测间隔设置
                if self.config.has_section('SETTINGS'):
                    self.detection_interval = self.config.getint('SETTINGS', 'detection_interval', fallback=5)
                else:
                    self.detection_interval = 5
            except Exception:
                self.detection_interval = 5
        else:
            self.detection_interval = 5
    
    def save_config(self):
        """保存配置文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
        except Exception:
            pass
    
    def save_detection_interval(self):
        """保存检测间隔设置"""
        try:
            interval = int(self.detection_interval_var.get())
            self.detection_interval = interval
            
            # 确保SETTINGS节存在
            if not self.config.has_section('SETTINGS'):
                self.config.add_section('SETTINGS')
            
            # 保存检测间隔
            self.config.set('SETTINGS', 'detection_interval', str(interval))
            
            # 保存到文件
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
        except Exception:
            pass
    
    def get_monitor_config_key(self, monitor):
        """获取显示器的配置键名"""
        # 使用显示器名称和分辨率作为唯一标识
        return f"{monitor['name']}_{monitor['width']}x{monitor['height']}"
    
    def save_monitor_config(self, monitor, auto_enabled, start_time, end_time):
        """保存显示器配置"""
        key = self.get_monitor_config_key(monitor)
        if key not in self.config:
            self.config.add_section(key)
        
        self.config[key]['auto_enabled'] = str(auto_enabled)
        self.config[key]['start_time'] = start_time
        self.config[key]['end_time'] = end_time
        self.config[key]['monitor_name'] = monitor['name']
        self.config[key]['resolution'] = f"{monitor['width']}x{monitor['height']}"
        
        self.save_config()
    
    def save_time_config(self, monitor):
        """保存时间配置变化"""
        try:
            # 找到对应的widget信息
            widget_info = None
            for widget in self.preview_widgets:
                if widget['monitor'] == monitor:
                    widget_info = widget
                    break
            
            if widget_info:
                start_time = f"{widget_info['start_hour_var'].get()}:{widget_info['start_min_var'].get()}:{widget_info['start_sec_var'].get()}"
                end_time = f"{widget_info['end_hour_var'].get()}:{widget_info['end_min_var'].get()}:{widget_info['end_sec_var'].get()}"
                self.save_monitor_config(monitor, widget_info['auto_enabled'], start_time, end_time)
        except Exception:
            pass
    
    def load_monitor_config(self, monitor):
        """加载显示器配置"""
        key = self.get_monitor_config_key(monitor)
        if key in self.config:
            try:
                return {
                    'auto_enabled': self.config.getboolean(key, 'auto_enabled', fallback=False),
                    'start_time': self.config.get(key, 'start_time', fallback='22:00:00'),
                    'end_time': self.config.get(key, 'end_time', fallback='08:00:00')
                }
            except Exception:
                pass
        return {
            'auto_enabled': False,
            'start_time': '17:0:0',
            'end_time': '7:0:0'
        }
    
    def get_screen_info(self):
        """获取所有屏幕信息"""
        try:
            monitors = get_monitors()
            screen_info = []
            
            # 获取显示器设备名称
            display_devices = self.get_display_devices()
            
            for i, monitor in enumerate(monitors):
                # 尝试获取真实的显示器名称
                device_name = f"显示器 {i + 1}"
                if i < len(display_devices):
                    device_name = display_devices[i].get('name', device_name)
                
                info = {
                    'id': i + 1,
                    'name': device_name,
                    'width': monitor.width,
                    'height': monitor.height,
                    'x': monitor.x,
                    'y': monitor.y,
                    'is_primary': monitor.is_primary
                }
                screen_info.append(info)
            return screen_info
        except Exception as e:
            messagebox.showerror("错误", f"获取屏幕信息失败: {str(e)}")
            return []
    
    def get_monitor_name_from_edid(self):
        """从注册表EDID信息中获取真实的显示器名称"""
        monitor_names = []
        try:
            # 打开显示器注册表项
            display_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                       r"SYSTEM\CurrentControlSet\Enum\DISPLAY")
            
            # 枚举所有显示器
            i = 0
            while True:
                try:
                    # 获取显示器子键名
                    monitor_key_name = winreg.EnumKey(display_key, i)
                    # 打开具体的显示器键
                    monitor_key = winreg.OpenKey(display_key, monitor_key_name)
                    
                    # 枚举显示器实例
                    j = 0
                    while True:
                        try:
                            instance_name = winreg.EnumKey(monitor_key, j)
                            
                            # 打开实例键
                            instance_key = winreg.OpenKey(monitor_key, instance_name)
                            
                            try:
                                # 尝试读取Device Parameters子键中的EDID
                                params_key = winreg.OpenKey(instance_key, "Device Parameters")
                                edid_data, _ = winreg.QueryValueEx(params_key, "EDID")
                                
                                # 解析EDID获取显示器名称
                                monitor_name = self.parse_edid_monitor_name(edid_data)
                                if monitor_name:
                                    monitor_names.append(monitor_name)
                                
                                winreg.CloseKey(params_key)
                            except FileNotFoundError:
                                pass  # 实例没有Device Parameters
                            except Exception:
                                pass  # 读取EDID失败
                            
                            winreg.CloseKey(instance_key)
                            j += 1
                        except OSError:
                            break
                    
                    winreg.CloseKey(monitor_key)
                    i += 1
                except OSError:
                    break
            
            winreg.CloseKey(display_key)
            
        except Exception:
            pass  # 注册表查询失败
        
        return monitor_names
    
    def parse_edid_monitor_name(self, edid_data):
        """从EDID数据中解析显示器名称"""
        try:
            if len(edid_data) < 128:
                return None
            
            # EDID中的显示器名称通常在字节54-125的描述符块中
            for i in range(54, 126, 18):  # 每个描述符块18字节
                if i + 17 < len(edid_data):
                    # 检查是否是显示器名称描述符 (类型 0xFC)
                    if edid_data[i + 3] == 0xFC:
                        # 提取名称 (从第5字节开始，最多13字节)
                        name_bytes = edid_data[i + 5:i + 18]
                        # 移除填充字符和空字符
                        name = name_bytes.rstrip(b'\x00\x0A\x20').decode('ascii', errors='ignore')
                        if name.strip():
                            return name.strip()
            
            return None
        except Exception:
            return None
    
    def get_display_devices(self):
        """获取显示设备信息"""
        devices = []
        try:
            # 首先尝试从注册表EDID获取真实的显示器名称
            monitor_devices = self.get_monitor_name_from_edid()
            
            # 如果EDID方法没有获取到足够的显示器，尝试使用WMI
            if not monitor_devices:
                try:
                    import wmi
                    c = wmi.WMI()
                    
                    # 首先尝试Win32_DesktopMonitor
                    try:
                        desktop_monitors = c.Win32_DesktopMonitor()
                        
                        for monitor in desktop_monitors:
                            if monitor.Name and monitor.Name.strip():
                                monitor_name = monitor.Name.strip()
                                # 过滤掉通用名称
                                if not any(generic in monitor_name.lower() for generic in ["default", "generic", "pnp", "plug and play", "默认", "通用", "即插即用"]):
                                    monitor_devices.append(monitor_name)
                    except Exception:
                        pass  # Win32_DesktopMonitor 查询失败
                        
                except ImportError:
                    pass  # WMI模块不可用
            
            # 如果还是没有获取到有效的显示器名称，使用screeninfo生成默认名称
            if not monitor_devices:
                from screeninfo import get_monitors
                monitors = get_monitors()
                monitor_devices = [f"显示器 {i + 1}" for i in range(len(monitors))]
            
            # 创建设备列表
            for i, name in enumerate(monitor_devices):
                devices.append({
                    'name': name,
                    'device_key': '',
                    'device_id': ''
                })
                    
        except ImportError:
                # 如果没有wmi模块，使用原来的方法
                i = 0
                while True:
                    try:
                        # 获取显示适配器信息
                        adapter = win32api.EnumDisplayDevices(None, i)
                        if adapter:
                            # 尝试获取连接到此适配器的显示器
                            monitor_name = f"显示器 {i + 1}"
                            
                            # 尝试枚举连接到此适配器的显示器
                            try:
                                monitor = win32api.EnumDisplayDevices(adapter.DeviceName, 0)
                                if monitor and monitor.DeviceString:
                                    # 清理显示器名称，移除不必要的前缀
                                    monitor_name = monitor.DeviceString
                                    # 移除常见的前缀
                                    prefixes_to_remove = ["Generic PnP Monitor", "即插即用监视器"]
                                    for prefix in prefixes_to_remove:
                                        if monitor_name.startswith(prefix):
                                            monitor_name = f"显示器 {i + 1}"
                                            break
                                    
                                    # 如果名称太长，截取合理长度
                                    if len(monitor_name) > 30:
                                        monitor_name = monitor_name[:27] + "..."
                            except:
                                pass
                            
                            devices.append({
                                'name': monitor_name,
                                'device_key': adapter.DeviceKey if hasattr(adapter, 'DeviceKey') else '',
                                'device_id': adapter.DeviceID if hasattr(adapter, 'DeviceID') else ''
                            })
                            i += 1
                        else:
                            break
                    except:
                        break
        except Exception as e:
            print(f"获取显示设备失败: {e}")
        
        return devices

    def get_device_name_by_monitor(self, monitor):
        """根据显示器信息获取设备名称"""
        try:
            # 枚举所有显示设备
            device_index = 0
            while True:
                try:
                    device = win32api.EnumDisplayDevices(None, device_index)
                    if not device:
                        break
                    
                    # 获取设备的显示设置
                    try:
                        settings = win32api.EnumDisplaySettings(device.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
                        # 比较位置和尺寸来匹配显示器
                        if (settings.Position_x == monitor['x'] and 
                            settings.Position_y == monitor['y'] and
                            settings.PelsWidth == monitor['width'] and
                            settings.PelsHeight == monitor['height']):
                            return device.DeviceName
                    except Exception as e:
                        # 如果获取当前设置失败（可能是熄屏状态），尝试获取注册表设置
                        try:
                            settings = win32api.EnumDisplaySettings(device.DeviceName, win32con.ENUM_REGISTRY_SETTINGS)
                            # 比较位置和尺寸来匹配显示器
                            if (settings.Position_x == monitor['x'] and 
                                settings.Position_y == monitor['y'] and
                                settings.PelsWidth == monitor['width'] and
                                settings.PelsHeight == monitor['height']):
                                return device.DeviceName
                        except Exception as e2:
                            # 如果都失败了，检查是否有保存的原始设置可以匹配
                            if hasattr(self, 'original_settings') and device.DeviceName in self.original_settings:
                                original = self.original_settings[device.DeviceName]
                                if (original['position_x'] == monitor['x'] and 
                                    original['position_y'] == monitor['y'] and
                                    original['width'] == monitor['width'] and
                                    original['height'] == monitor['height']):
                                    return device.DeviceName
                            # 静默处理错误，避免在熄屏状态下产生错误信息
                            pass
                    
                    device_index += 1
                except:
                    break
            
            # 如果没有找到精确匹配，尝试按索引匹配
            # 假设monitor['id']对应设备索引
            try:
                device = win32api.EnumDisplayDevices(None, monitor['id'] - 1)
                if device:
                    return device.DeviceName
            except:
                pass
                
            return None
        except Exception as e:
            return None

    def create_widgets(self):
        """创建主界面组件"""
        # 预览区域框架
        preview_frame = tk.Frame(self.root)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 确保至少显示两个屏幕预览（即使只有一个屏幕）
        display_monitors = self.monitors[:2] if len(self.monitors) >= 2 else self.monitors + [None] * (2 - len(self.monitors))
        
        self.preview_widgets = []
        
        for i, monitor in enumerate(display_monitors):
            # 每个屏幕的容器
            screen_frame = tk.Frame(preview_frame, relief=tk.RAISED, borderwidth=2)
            screen_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
            
            if monitor:
                # 显示器名称（移到预览窗口上方）
                title = f"{monitor['name']} ({'主屏' if monitor['is_primary'] else '副屏'})"
                screen_title = tk.Label(screen_frame, text=title, 
                                      font=("./29华康宋体W3.ttf", 12, "bold"))
                screen_title.pack(pady=(5, 2))
                
                # 预览窗口
                preview_canvas = tk.Canvas(screen_frame, width=300, height=200, 
                                         bg="black", relief=tk.SUNKEN, borderwidth=2)
                preview_canvas.pack(pady=2)
                
                # 分辨率信息（移到预览窗口下方）
                info_text = f"分辨率: {monitor['width']} x {monitor['height']}"
                info_label = tk.Label(screen_frame, text=info_text, 
                                    font=("./29华康宋体W3.ttf", 10))
                info_label.pack(pady=2)
                
                # 自动熄屏时间设置框架
                auto_frame = tk.Frame(screen_frame)
                auto_frame.pack(pady=2)
                
                # 加载显示器配置
                config = self.load_monitor_config(monitor)
                start_parts = config['start_time'].split(':')
                end_parts = config['end_time'].split(':')
                
                # 开始时间设置
                start_time_frame = tk.Frame(auto_frame)
                start_time_frame.pack(pady=1)
                tk.Label(start_time_frame, text="开始时间:", font=("./29华康宋体W3.ttf", 8)).pack(side=tk.LEFT)
                start_hour_var = tk.StringVar(value=start_parts[0] if len(start_parts) > 0 else "18")
                start_min_var = tk.StringVar(value=start_parts[1] if len(start_parts) > 1 else "00")
                start_sec_var = tk.StringVar(value=start_parts[2] if len(start_parts) > 2 else "00")
                # 创建时间变化回调函数
                def on_time_change(*args, m=monitor):
                    self.save_time_config(m)
                
                start_hour_var.trace('w', on_time_change)
                start_min_var.trace('w', on_time_change)
                start_sec_var.trace('w', on_time_change)
                
                # 创建开始时间的Spinbox控件
                start_hour_spinbox = tk.Spinbox(start_time_frame, from_=0, to=23, width=3, textvariable=start_hour_var, 
                          font=("./29华康宋体W3.ttf", 8), state="disabled" if config['auto_enabled'] else "normal")
                start_hour_spinbox.pack(side=tk.LEFT, padx=1)
                tk.Label(start_time_frame, text=":", font=("./29华康宋体W3.ttf", 8)).pack(side=tk.LEFT)
                start_min_spinbox = tk.Spinbox(start_time_frame, from_=0, to=59, width=3, textvariable=start_min_var, 
                          font=("./29华康宋体W3.ttf", 8), state="disabled" if config['auto_enabled'] else "normal")
                start_min_spinbox.pack(side=tk.LEFT, padx=1)
                tk.Label(start_time_frame, text=":", font=("./29华康宋体W3.ttf", 8)).pack(side=tk.LEFT)
                start_sec_spinbox = tk.Spinbox(start_time_frame, from_=0, to=59, width=3, textvariable=start_sec_var, 
                          font=("./29华康宋体W3.ttf", 8), state="disabled" if config['auto_enabled'] else "normal")
                start_sec_spinbox.pack(side=tk.LEFT, padx=1)
                
                # 结束时间设置
                end_time_frame = tk.Frame(auto_frame)
                end_time_frame.pack(pady=1)
                tk.Label(end_time_frame, text="结束时间:", font=("./29华康宋体W3.ttf", 8)).pack(side=tk.LEFT)
                end_hour_var = tk.StringVar(value=end_parts[0] if len(end_parts) > 0 else "07")
                end_min_var = tk.StringVar(value=end_parts[1] if len(end_parts) > 1 else "00")
                end_sec_var = tk.StringVar(value=end_parts[2] if len(end_parts) > 2 else "00")
                end_hour_var.trace('w', on_time_change)
                end_min_var.trace('w', on_time_change)
                end_sec_var.trace('w', on_time_change)
                
                # 创建结束时间的Spinbox控件
                end_hour_spinbox = tk.Spinbox(end_time_frame, from_=0, to=23, width=3, textvariable=end_hour_var, 
                          font=("./29华康宋体W3.ttf", 8), state="disabled" if config['auto_enabled'] else "normal")
                end_hour_spinbox.pack(side=tk.LEFT, padx=1)
                tk.Label(end_time_frame, text=":", font=("./29华康宋体W3.ttf", 8)).pack(side=tk.LEFT)
                end_min_spinbox = tk.Spinbox(end_time_frame, from_=0, to=59, width=3, textvariable=end_min_var, 
                          font=("./29华康宋体W3.ttf", 8), state="disabled" if config['auto_enabled'] else "normal")
                end_min_spinbox.pack(side=tk.LEFT, padx=1)
                tk.Label(end_time_frame, text=":", font=("./29华康宋体W3.ttf", 8)).pack(side=tk.LEFT)
                end_sec_spinbox = tk.Spinbox(end_time_frame, from_=0, to=59, width=3, textvariable=end_sec_var, 
                          font=("./29华康宋体W3.ttf", 8), state="disabled" if config['auto_enabled'] else "normal")
                end_sec_spinbox.pack(side=tk.LEFT, padx=1)
                
                # 按钮框架
                button_frame = tk.Frame(screen_frame)
                button_frame.pack(pady=2)
                
                # 手动熄屏按钮
                manual_off_btn = tk.Button(button_frame, text="手动熄屏", 
                                         command=lambda m=monitor: self.turn_off_screen(m),
                                         bg="#ff6b6b", fg="white", 
                                         font=("./29华康宋体W3.ttf", 9, "bold"),
                                         width=8, height=1)
                manual_off_btn.pack(side=tk.LEFT, padx=2)
                
                # 自动熄屏按钮
                auto_off_btn = tk.Button(button_frame, text="自动熄屏", 
                                       command=lambda m=monitor: self.toggle_auto_screen_off(m),
                                       bg="#fd7e14", fg="white", 
                                       font=("./29华康宋体W3.ttf", 9, "bold"),
                                       width=8, height=1)
                auto_off_btn.pack(side=tk.LEFT, padx=2)
                
                # 自动熄屏状态显示
                auto_status_text = "启用" if config['auto_enabled'] else "禁用"
                auto_status_color = "green" if config['auto_enabled'] else "red"
                auto_status_label = tk.Label(button_frame, text=auto_status_text, 
                                            font=("./29华康宋体W3.ttf", 8),
                                            fg=auto_status_color)
                auto_status_label.pack(side=tk.LEFT, padx=2)
                

                
                self.preview_widgets.append({
                    'canvas': preview_canvas,
                    'monitor': monitor,
                    'title': screen_title,
                    'info': info_label,
                    'start_hour_var': start_hour_var,
                    'start_min_var': start_min_var,
                    'start_sec_var': start_sec_var,
                    'end_hour_var': end_hour_var,
                    'end_min_var': end_min_var,
                    'end_sec_var': end_sec_var,
                    'start_hour_spinbox': start_hour_spinbox,
                    'start_min_spinbox': start_min_spinbox,
                    'start_sec_spinbox': start_sec_spinbox,
                    'end_hour_spinbox': end_hour_spinbox,
                    'end_min_spinbox': end_min_spinbox,
                    'end_sec_spinbox': end_sec_spinbox,
                    'auto_status_label': auto_status_label,
                    'auto_enabled': config['auto_enabled']
                })
            else:
                # 无屏幕时的占位符
                no_screen_label = tk.Label(screen_frame, text="未检测到屏幕", 
                                         font=("./29华康宋体W3.ttf", 12),
                                         fg="gray")
                no_screen_label.pack(expand=True)
        
        # 底部控制区域
        control_frame = tk.Frame(self.root)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 刷新按钮
        refresh_btn = tk.Button(control_frame, text="刷新屏幕信息", 
                              command=self.refresh_screens,
                              bg="#339af0", fg="white", 
                              font=("./29华康宋体W3.ttf", 10, "bold"))
        refresh_btn.pack(side=tk.LEFT, padx=5)
        
        # 检测间隔设置
        interval_frame = tk.Frame(control_frame)
        interval_frame.pack(side=tk.LEFT, padx=10)
        
        tk.Label(interval_frame, text="检测间隔:", font=("./29华康宋体W3.ttf", 9)).pack(side=tk.LEFT)
        
        self.detection_interval_var = tk.StringVar(value=str(self.detection_interval))
        interval_spinbox = tk.Spinbox(interval_frame, from_=1, to=60, width=5,
                                    textvariable=self.detection_interval_var,
                                    font=("./29华康宋体W3.ttf", 9),
                                    command=self.save_detection_interval)
        interval_spinbox.pack(side=tk.LEFT, padx=2)
        
        tk.Label(interval_frame, text="秒", font=("./29华康宋体W3.ttf", 9)).pack(side=tk.LEFT)
        
        # 绑定变量变化事件
        self.detection_interval_var.trace('w', lambda *args: self.save_detection_interval())
        
        # 显示模式控制区域（中间）
        display_mode_frame = tk.Frame(control_frame)
        display_mode_frame.pack(side=tk.LEFT, expand=True)
        
        # 退出按钮
        exit_btn = tk.Button(control_frame, text="退出", 
                           command=self.on_closing,
                           bg="#e03131", fg="white", 
                           font=("./29华康宋体W3.ttf", 10, "bold"))
        exit_btn.pack(side=tk.RIGHT, padx=5)
        
        # 重置显示器按钮
        reset_btn = tk.Button(control_frame, text="重置显示器", 
                            command=self.reset_displays,
                            bg="#fd7e14", fg="white", 
                            font=("./29华康宋体W3.ttf", 10, "bold"),
                            width=12)
        reset_btn.pack(side=tk.RIGHT, padx=5)
        
        # 重置显示器说明文本
        reset_info_label = tk.Label(control_frame, text="熄屏后需用重置显示器恢复。也可用快捷键CTRL+ALT+X恢复⇢", 
                                   font=("./29华康宋体W3.ttf", 9), fg="#666666")
        reset_info_label.pack(side=tk.RIGHT, padx=5)
        
        # 为重置按钮添加工具提示
        ToolTip(reset_btn, "快捷键: Ctrl+Alt+X")
    
    def capture_screen_preview(self, monitor):
        """捕获屏幕预览"""
        try:
            # 获取指定屏幕区域的截图
            left = monitor['x']
            top = monitor['y']
            right = left + monitor['width']
            bottom = top + monitor['height']
            
            # 使用PIL的ImageGrab捕获整个虚拟屏幕，然后裁剪指定区域
            # 这样可以处理负坐标的情况
            # 计算所有屏幕的边界来确定虚拟屏幕范围
            all_left = min(m['x'] for m in self.monitors)
            all_top = min(m['y'] for m in self.monitors)
            all_right = max(m['x'] + m['width'] for m in self.monitors)
            all_bottom = max(m['y'] + m['height'] for m in self.monitors)
            
            # 捕获包含所有屏幕的区域
            # 使用all_screens=True参数来支持多屏幕截图
            full_screenshot = ImageGrab.grab(bbox=(all_left, all_top, all_right, all_bottom), all_screens=True)
            
            # 获取虚拟屏幕的总尺寸
            virtual_width = full_screenshot.width
            virtual_height = full_screenshot.height
            
            # 计算在捕获图像中的相对坐标
            # 由于我们捕获的是从(all_left, all_top)开始的区域
            crop_left = left - all_left
            crop_top = top - all_top
            crop_right = crop_left + monitor['width']
            crop_bottom = crop_top + monitor['height']
            
            # 确保裁剪区域在有效范围内
            crop_left = max(0, crop_left)
            crop_top = max(0, crop_top)
            crop_right = min(virtual_width, crop_right)
            crop_bottom = min(virtual_height, crop_bottom)
            
            # 裁剪指定区域
            screenshot = full_screenshot.crop((crop_left, crop_top, crop_right, crop_bottom))
            
            # 缩放到预览窗口大小
            preview_width, preview_height = 300, 200
            screenshot = screenshot.resize((preview_width, preview_height), Image.Resampling.LANCZOS)
            
            # 直接返回截图，不添加任何文字覆盖层
            return ImageTk.PhotoImage(screenshot)
            
        except Exception as e:
            # 如果截图失败，创建一个错误提示图像
            return self.create_error_preview(monitor, str(e))
    
    def create_error_preview(self, monitor, error_msg):
        """创建错误提示预览图像"""
        try:
            width, height = 300, 200
            image = Image.new('RGB', (width, height), color='#2c3e50')
            
            from PIL import ImageDraw, ImageFont
            draw = ImageDraw.Draw(image)
            
            try:
                font = ImageFont.truetype("./29华康宋体W3.ttf", 12)
            except:
                font = ImageFont.load_default()
            
            # 绘制错误信息
            text = f"{monitor['name']}\n{monitor['width']} x {monitor['height']}\n\n无法获取预览\n{error_msg[:30]}..."
            draw.text((10, 10), text, fill='white', font=font)
            
            return ImageTk.PhotoImage(image)
        except:
            return None
    
    def update_previews(self):
        """更新屏幕预览"""
        while self.preview_running:
            try:
                # 检查preview_widgets是否存在且不为空
                if not hasattr(self, 'preview_widgets') or not self.preview_widgets:
                    time.sleep(1)
                    continue
                    
                for i, widget_info in enumerate(self.preview_widgets):
                    if not self.preview_running:
                        break
                        
                    if widget_info['monitor']:
                        # 检查canvas是否仍然有效
                        try:
                            canvas = widget_info['canvas']
                            canvas.winfo_exists()
                        except Exception:
                            continue
                        
                        preview_image = self.capture_screen_preview(widget_info['monitor'])
                        if preview_image:
                            try:
                                canvas.delete("all")
                                canvas.create_image(150, 100, image=preview_image)
                                canvas.image = preview_image  # 保持引用
                            except Exception:
                                pass
                
                time.sleep(1)  # 每秒更新一次
            except Exception:
                time.sleep(1)
    
    def turn_off_screen(self, monitor):
        """熄灭指定屏幕"""
        try:
            # 获取显示设备信息
            device_name = self.get_device_name_by_monitor(monitor)
            if not device_name:
                messagebox.showerror("错误", f"无法找到显示器 {monitor['name']} 的设备名称")
                return
            
            # 获取当前显示设置
            current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
            
            # 保存原始设置用于恢复
            if not hasattr(self, 'original_settings'):
                self.original_settings = {}
            self.original_settings[device_name] = {
                'width': current_settings.PelsWidth,
                'height': current_settings.PelsHeight,
                'position_x': current_settings.Position_x,
                'position_y': current_settings.Position_y
            }
            
            # 创建新的显示设置（关闭显示器）
            new_settings = current_settings
            new_settings.PelsWidth = 0
            new_settings.PelsHeight = 0
            new_settings.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT | win32con.DM_POSITION
            
            # 应用设置
            result = win32api.ChangeDisplaySettingsEx(device_name, new_settings, 0)
            if result != win32con.DISP_CHANGE_SUCCESSFUL:
                messagebox.showerror("错误", f"熄屏失败，错误代码: {result}")
                
        except Exception as e:
            messagebox.showerror("错误", f"熄屏失败: {str(e)}")
    

    

    

    
    def force_refresh_displays(self):
        """强制刷新显示器配置"""
        try:
            # 方法1: 使用ChangeDisplaySettings刷新所有显示器
            result = win32api.ChangeDisplaySettings(None, 0)
            if result == win32con.DISP_CHANGE_SUCCESSFUL:
                print("成功刷新显示器配置")
                return True
            else:
                print(f"刷新显示器配置失败，错误代码: {result}")
                
            # 方法2: 如果方法1失败，尝试重新应用当前设置
            import time
            time.sleep(0.2)
            
            # 枚举所有显示设备并重新应用设置
            i = 0
            refreshed_count = 0
            while True:
                try:
                    device = win32api.EnumDisplayDevices(None, i)
                    if not device.DeviceName:
                        break
                        
                    device_name = device.DeviceName
                    
                    # 获取当前设置并重新应用
                    try:
                        current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
                        if current_settings.PelsWidth > 0 and current_settings.PelsHeight > 0:
                            result = win32api.ChangeDisplaySettingsEx(device_name, current_settings, win32con.CDS_UPDATEREGISTRY)
                            if result == win32con.DISP_CHANGE_SUCCESSFUL:
                                refreshed_count += 1
                    except Exception:
                        pass
                    
                    i += 1
                except:
                    break
            
            if refreshed_count > 0:
                return True
                
            return False
            
        except Exception:
            return False
    
    def restore_all_screens(self):
        """恢复所有被熄屏的显示器"""
        try:
            if not hasattr(self, 'original_settings') or not self.original_settings:
                return
            
            # 首先尝试强制刷新显示器配置
            if self.force_refresh_displays():
                # 等待系统处理
                import time
                time.sleep(1.0)
                
                # 清空原始设置，因为已经通过刷新恢复了
                self.original_settings.clear()
                return
            
            # 重新枚举所有显示设备
            available_devices = []
            i = 0
            while True:
                try:
                    device = win32api.EnumDisplayDevices(None, i)
                    if device.DeviceName:
                        available_devices.append(device.DeviceName)
                    i += 1
                except:
                    break
            
            restored_count = 0
            for device_name, original in self.original_settings.items():
                try:
                    # 检查设备是否仍然存在
                    if device_name not in available_devices:
                        continue
                    
                    # 尝试获取当前设置
                    try:
                        current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
                    except:
                        # 如果获取当前设置失败，尝试获取注册表设置
                        try:
                            current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_REGISTRY_SETTINGS)
                        except:
                            continue
                    
                    # 创建新的设置
                    new_settings = current_settings
                    new_settings.PelsWidth = original['width']
                    new_settings.PelsHeight = original['height']
                    new_settings.Position_x = original['position_x']
                    new_settings.Position_y = original['position_y']
                    new_settings.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT | win32con.DM_POSITION
                    
                    result = win32api.ChangeDisplaySettingsEx(device_name, new_settings, win32con.CDS_UPDATEREGISTRY)
                    if result == win32con.DISP_CHANGE_SUCCESSFUL:
                        restored_count += 1
                        
                except Exception:
                    pass
            
            if restored_count > 0:
                # 清空已恢复的设置
                self.original_settings.clear()
                
        except Exception:
            pass
    
    def get_all_available_monitors(self):
        """获取所有可用的显示器，包括被熄屏的"""
        try:
            # 获取当前检测到的显示器
            current_monitors = self.get_screen_info()
            
            # 如果有保存的原始设置，说明有显示器被熄屏了
            if hasattr(self, 'original_settings') and self.original_settings:
                # 检查是否有被熄屏但未在当前列表中的显示器
                current_device_names = set()
                for monitor in current_monitors:
                    device_name = self.get_device_name_by_monitor(monitor)
                    if device_name:
                        current_device_names.add(device_name)
                
                # 统计总的可用显示器数量（包括被熄屏的）
                total_available = len(current_monitors) + len([name for name in self.original_settings.keys() if name not in current_device_names])
                return total_available, current_monitors
            
            return len(current_monitors), current_monitors
        except Exception:
            return len(self.monitors), self.monitors
    
    def set_duplicate_mode_api(self):
        """使用API设置复制模式"""
        try:
            # 智能检测显示器数量，包括被熄屏的
            total_monitors, current_monitors = self.get_all_available_monitors()
            
            if total_monitors < 2:
                messagebox.showwarning("警告", "需要至少两个显示器才能使用复制模式")
                return
            
            # 首先恢复所有被熄屏的显示器
            self.restore_all_screens()
            
            # 等待系统处理恢复操作
            import time
            time.sleep(0.5)
            
            # 重新获取显示器信息
            self.monitors = self.get_screen_info()
            
            # 获取主显示器信息
            primary_monitor = None
            for monitor in self.monitors:
                if monitor['is_primary']:
                    primary_monitor = monitor
                    break
            
            if not primary_monitor:
                messagebox.showerror("错误", "未找到主显示器")
                return
            
            # 将所有显示器设置为与主显示器相同的位置和分辨率
            for monitor in self.monitors:
                if not monitor['is_primary']:
                    device_name = self.get_device_name_by_monitor(monitor)
                    if device_name:
                        try:
                            current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
                            new_settings = current_settings
                            new_settings.Position_x = 0
                            new_settings.Position_y = 0
                            new_settings.PelsWidth = primary_monitor['width']
                            new_settings.PelsHeight = primary_monitor['height']
                            new_settings.Fields = win32con.DM_POSITION | win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT
                            
                            result = win32api.ChangeDisplaySettingsEx(device_name, new_settings, 0)
                        except Exception:
                            pass
            
            # 刷新屏幕信息
            self.root.after(1000, self.refresh_screens)
            
        except Exception as e:
            messagebox.showerror("错误", f"API切换复制模式失败: {str(e)}")
    
    def set_extend_mode(self):
        """设置扩展模式（对应Win+P扩展）"""
        try:
            # 直接使用API方法，避免调出Win+P菜单
            self.set_extend_mode_api()
                
        except Exception as e:
            messagebox.showerror("错误", f"切换扩展模式失败: {str(e)}")
    
    def set_extend_mode_api(self):
        """使用API设置扩展模式"""
        try:
            # 智能检测显示器数量，包括被熄屏的
            total_monitors, current_monitors = self.get_all_available_monitors()
            
            if total_monitors < 2:
                messagebox.showwarning("警告", "需要至少两个显示器才能使用扩展模式")
                return
            
            # 首先恢复所有被熄屏的显示器
            self.restore_all_screens()
            
            # 等待系统处理恢复操作
            import time
            time.sleep(0.5)
            
            # 重新获取显示器信息，确保使用最新状态
            self.monitors = self.get_screen_info()
            
            # 重新排列显示器位置以实现扩展模式
            x_offset = 0
            primary_width = 0
            
            # 首先处理主显示器
            for monitor in self.monitors:
                if monitor['is_primary']:
                    device_name = self.get_device_name_by_monitor(monitor)
                    if device_name:
                        try:
                            current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
                            new_settings = current_settings
                            
                            # 确保主显示器有正确的分辨率和位置
                            if hasattr(self, 'original_settings') and device_name in self.original_settings:
                                original = self.original_settings[device_name]
                                new_settings.PelsWidth = original['width']
                                new_settings.PelsHeight = original['height']
                                primary_width = original['width']
                            else:
                                new_settings.PelsWidth = monitor['width']
                                new_settings.PelsHeight = monitor['height']
                                primary_width = monitor['width']
                            
                            new_settings.Position_x = 0
                            new_settings.Position_y = 0
                            new_settings.Fields = win32con.DM_POSITION | win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT
                            
                            result = win32api.ChangeDisplaySettingsEx(device_name, new_settings, 0)
                        except Exception:
                            pass
                    break
            
            # 然后处理副显示器
            x_offset = primary_width
            for monitor in self.monitors:
                if not monitor['is_primary']:
                    device_name = self.get_device_name_by_monitor(monitor)
                    if device_name:
                        try:
                            current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
                            new_settings = current_settings
                            
                            # 确保副显示器有正确的分辨率
                            if hasattr(self, 'original_settings') and device_name in self.original_settings:
                                original = self.original_settings[device_name]
                                new_settings.PelsWidth = original['width']
                                new_settings.PelsHeight = original['height']
                            else:
                                new_settings.PelsWidth = monitor['width']
                                new_settings.PelsHeight = monitor['height']
                            
                            # 副显示器放在主显示器右侧
                            new_settings.Position_x = x_offset
                            new_settings.Position_y = 0
                            new_settings.Fields = win32con.DM_POSITION | win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT
                            
                            result = win32api.ChangeDisplaySettingsEx(device_name, new_settings, 0)
                            if result == win32con.DISP_CHANGE_SUCCESSFUL:
                                x_offset += new_settings.PelsWidth
                        except Exception:
                            pass
            
            # 刷新屏幕信息
            self.root.after(1000, self.refresh_screens)
            
        except Exception as e:
            messagebox.showerror("错误", f"API切换扩展模式失败: {str(e)}")
    

    
    def reset_displays(self):
        """重置显示器 - 只重新加载显示器，不改变排列数据"""
        try:
            # 使用强制刷新显示器的方法，这不会改变显示器的排列
            if self.force_refresh_displays():
                # 等待系统处理
                import time
                time.sleep(1.0)
                
                # 清空原始设置，因为已经通过刷新恢复了
                if hasattr(self, 'original_settings'):
                    self.original_settings.clear()
                
                # 刷新屏幕信息以更新UI
                self.refresh_screens()
                
        except Exception:
            pass
    
    def start_auto_screen_timer(self):
        """启动自动熄屏检测定时器"""
        def timer_loop():
            while self.auto_screen_timer_running:
                try:
                    self.check_auto_screen_off()
                    time.sleep(self.detection_interval)  # 使用配置的检测间隔
                except Exception:
                    time.sleep(self.detection_interval)
        
        timer_thread = threading.Thread(target=timer_loop, daemon=True)
        timer_thread.start()
    
    def check_auto_screen_off(self):
        """检测并执行自动熄屏逻辑"""
        try:
            from datetime import datetime
            
            # 检查preview_widgets是否存在
            if not hasattr(self, 'preview_widgets') or not self.preview_widgets:
                return
            
            current_time = datetime.now()
            current_hour = current_time.hour
            current_min = current_time.minute
            current_sec = current_time.second
            
            for widget_info in self.preview_widgets:
                try:
                    # 检查widget_info是否有效
                    if not isinstance(widget_info, dict) or not widget_info.get('auto_enabled', False):
                        continue
                    
                    # 安全地获取时间设置变量
                    start_hour_var = widget_info.get('start_hour_var')
                    start_min_var = widget_info.get('start_min_var')
                    start_sec_var = widget_info.get('start_sec_var')
                    end_hour_var = widget_info.get('end_hour_var')
                    end_min_var = widget_info.get('end_min_var')
                    end_sec_var = widget_info.get('end_sec_var')
                    
                    # 检查所有变量是否存在
                    if not all([start_hour_var, start_min_var, start_sec_var, end_hour_var, end_min_var, end_sec_var]):
                        continue
                    
                    # 获取设置的时间范围
                    start_hour = int(start_hour_var.get())
                    start_min = int(start_min_var.get())
                    start_sec = int(start_sec_var.get())
                    end_hour = int(end_hour_var.get())
                    end_min = int(end_min_var.get())
                    end_sec = int(end_sec_var.get())
                    
                    # 转换为秒数进行比较
                    current_total_sec = current_hour * 3600 + current_min * 60 + current_sec
                    start_total_sec = start_hour * 3600 + start_min * 60 + start_sec
                    end_total_sec = end_hour * 3600 + end_min * 60 + end_sec
                    
                    # 判断是否在熄屏时间范围内
                    in_sleep_time = False
                    if start_total_sec <= end_total_sec:
                        # 同一天内的时间范围
                        in_sleep_time = start_total_sec <= current_total_sec <= end_total_sec
                    else:
                        # 跨天的时间范围（如18:00到次日07:00）
                        in_sleep_time = current_total_sec >= start_total_sec or current_total_sec <= end_total_sec
                    
                    monitor = widget_info.get('monitor')
                    if not monitor:
                        continue
                    
                    if in_sleep_time:
                        # 在熄屏时间范围内，检测屏幕是否已熄屏，如果没有则熄屏
                        if self.is_monitor_on(monitor):
                            self.turn_off_screen(monitor)
                    else:
                        # 不在熄屏时间范围内，检测屏幕是否熄屏，如果熄屏则开启
                        if not self.is_monitor_on(monitor):
                            self.reset_displays()
                            
                except Exception:
                    # 单个显示器检测失败时继续处理其他显示器
                    continue
                    
        except Exception:
            # 整个检测过程失败时静默处理
            pass
    
    def is_monitor_on(self, monitor):
        """检测显示器是否开启"""
        try:
            # 获取显示器对应的设备名称
            device_name = self.get_device_name_by_monitor(monitor)
            if not device_name:
                return True  # 如果无法获取设备名称，默认认为是开启的
            
            # 检查是否有保存的原始设置（表示该显示器被熄屏了）
            if hasattr(self, 'original_settings') and device_name in self.original_settings:
                return False  # 如果有原始设置记录，说明显示器被熄屏了
            
            # 尝试获取当前显示设置来判断显示器状态
            try:
                current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
                # 如果能成功获取设置且分辨率不为0，说明显示器是开启的
                return current_settings.PelsWidth > 0 and current_settings.PelsHeight > 0
            except Exception:
                # 如果无法获取当前设置，可能是显示器被熄屏了
                return False
                
        except Exception:
            return True  # 出错时默认认为是开启的
    
    def toggle_auto_screen_off(self, monitor):
        """切换自动熄屏状态"""
        try:
            # 找到对应的widget信息
            widget_info = None
            for widget in self.preview_widgets:
                if widget['monitor'] == monitor:
                    widget_info = widget
                    break
            
            if not widget_info:
                return
            
            # 切换自动熄屏状态
            widget_info['auto_enabled'] = not widget_info['auto_enabled']
            
            # 更新状态显示和时间输入框状态
            if widget_info['auto_enabled']:
                widget_info['auto_status_label'].config(text="启用", fg="green")
                # 启用自动熄屏时，禁用时间输入框
                widget_info['start_hour_spinbox'].config(state="disabled")
                widget_info['start_min_spinbox'].config(state="disabled")
                widget_info['start_sec_spinbox'].config(state="disabled")
                widget_info['end_hour_spinbox'].config(state="disabled")
                widget_info['end_min_spinbox'].config(state="disabled")
                widget_info['end_sec_spinbox'].config(state="disabled")
                start_time = f"{widget_info['start_hour_var'].get()}:{widget_info['start_min_var'].get()}:{widget_info['start_sec_var'].get()}"
                end_time = f"{widget_info['end_hour_var'].get()}:{widget_info['end_min_var'].get()}:{widget_info['end_sec_var'].get()}"
            else:
                widget_info['auto_status_label'].config(text="禁用", fg="red")
                # 禁用自动熄屏时，启用时间输入框
                widget_info['start_hour_spinbox'].config(state="normal")
                widget_info['start_min_spinbox'].config(state="normal")
                widget_info['start_sec_spinbox'].config(state="normal")
                widget_info['end_hour_spinbox'].config(state="normal")
                widget_info['end_min_spinbox'].config(state="normal")
                widget_info['end_sec_spinbox'].config(state="normal")
                start_time = f"{widget_info['start_hour_var'].get()}:{widget_info['start_min_var'].get()}:{widget_info['start_sec_var'].get()}"
                end_time = f"{widget_info['end_hour_var'].get()}:{widget_info['end_min_var'].get()}:{widget_info['end_sec_var'].get()}"
            
            # 保存配置到INI文件
            self.save_monitor_config(monitor, widget_info['auto_enabled'], start_time, end_time)
                
        except Exception:
            pass
    
    def refresh_screens(self):
        """刷新屏幕信息"""
        # 停止预览更新线程
        self.preview_running = False
        if hasattr(self, 'preview_thread') and self.preview_thread.is_alive():
            self.preview_thread.join(timeout=2)
        
        # 清空预览组件列表，防止线程访问已销毁的组件
        self.preview_widgets.clear()
        
        self.monitors = self.get_screen_info()
        # 重新创建界面
        for widget in self.root.winfo_children():
            widget.destroy()
        self.create_widgets()
        
        # 重新启动预览更新线程
        self.preview_running = True
        self.preview_thread = threading.Thread(target=self.update_previews, daemon=True)
        self.preview_thread.start()
    
    def run(self):
        """启动应用程序"""
        try:
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            self.root.mainloop()
        except Exception as e:
            messagebox.showerror("错误", f"程序运行错误: {str(e)}")
    
    def setup_global_hotkeys(self):
        """设置全局快捷键"""
        try:
            # 注册CTRL+ALT+X快捷键
            keyboard.add_hotkey('ctrl+alt+x', self.hotkey_reset_displays)
        except Exception as e:
            print(f"注册全局快捷键失败: {str(e)}")
    
    def hotkey_reset_displays(self):
        """快捷键触发的重置显示器方法"""
        try:
            # 在主线程中执行重置显示器操作
            self.root.after(0, self.reset_displays)
        except Exception as e:
            print(f"快捷键重置显示器失败: {str(e)}")
    
    def on_closing(self):
        """程序关闭时的清理工作"""
        self.preview_running = False
        self.auto_screen_timer_running = False
        # 清理全局快捷键
        try:
            keyboard.unhook_all_hotkeys()
        except:
            pass
        self.root.destroy()

# 全局变量保存socket对象
_single_instance_socket = None

def check_single_instance():
    """检查是否已有程序实例在运行"""
    import socket
    import atexit
    global _single_instance_socket
    
    # 使用特定端口作为单实例标识
    SINGLE_INSTANCE_PORT = 19876
    
    try:
        # 创建socket并绑定端口
        _single_instance_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        _single_instance_socket.bind(('127.0.0.1', SINGLE_INSTANCE_PORT))
        _single_instance_socket.listen(1)
        
        # 注册退出时清理socket
        def cleanup_socket():
            global _single_instance_socket
            if _single_instance_socket:
                try:
                    _single_instance_socket.close()
                except:
                    pass
                _single_instance_socket = None
        
        atexit.register(cleanup_socket)
        
        print("单实例检查通过，程序正常启动")
        return True
        
    except OSError:
        # 端口已被占用，说明已有实例运行
        print("检测到已有程序实例运行，正在激活现有窗口...")
        activate_existing_window()
        return False
    except Exception as e:
        print(f"单实例检查出现异常: {e}")
        # 如果检查失败，允许程序继续运行
        return True

def activate_existing_window():
    """激活已存在的程序窗口"""
    try:
        # 尝试使用pywin32激活窗口
        import win32gui
        import win32con
        
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if "老绅控屏眼" in window_text:
                    windows.append(hwnd)
            return True
        
        # 查找所有窗口
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        # 激活找到的窗口
        for hwnd in windows:
            # 恢复窗口（如果最小化）
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            # 将窗口置前
            win32gui.SetForegroundWindow(hwnd)
            break
            
        print("已尝试激活现有窗口")
        
    except ImportError:
        # pywin32不可用时的备用方案
        print("pywin32不可用，无法激活现有窗口")
    except Exception as e:
        print(f"激活窗口时出现异常: {e}")

def main():
    """主函数"""
    # 检查单实例
    if not check_single_instance():
        return  # 已有实例运行，退出
    
    try:
        app = ScreenController()
        app.run()
    except Exception:
        input("按回车键退出...")

if __name__ == "__main__":
    main()