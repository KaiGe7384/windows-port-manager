import customtkinter as ctk
import subprocess
import threading
import tkinter.messagebox as messagebox
from tkinter import ttk
import os
import sys
import warnings

# 抑制PIL/PNG警告
warnings.filterwarnings("ignore", category=UserWarning, module="PIL")
os.environ["PYTHONWARNINGS"] = "ignore"

# 设置外观模式和颜色主题
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class PortManagerApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Windows 端口管理工具")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        
        # 创建界面
        self.create_widgets()
        
        # 检查管理员权限（在界面创建后）
        self.check_admin_rights()
        
    def check_admin_rights(self):
        """检查是否以管理员身份运行"""
        try:
            import ctypes
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if not is_admin:
                response = messagebox.askyesno(
                    "权限不足", 
                    "检测到程序未以管理员身份运行！\n\n" +
                    "端口管理和防火墙功能需要管理员权限才能正常工作。\n\n" +
                    "是否重新以管理员身份启动程序？\n\n" +
                    "选择'是'将关闭当前程序并请求管理员权限\n" +
                    "选择'否'将继续运行但功能受限"
                )
                if response:
                    self.restart_as_admin()
                    return False
                else:
                    self.update_status("⚠️ 警告：未以管理员身份运行，部分功能可能无法正常工作")
            else:
                self.update_status("✅ 已检测到管理员权限，所有功能可正常使用")
            return True
        except Exception as e:
            self.update_status(f"❌ 权限检查失败: {e}")
            return True
    
    def restart_as_admin(self):
        """重新以管理员身份启动程序"""
        try:
            import sys
            import ctypes
            if sys.argv[0].endswith('.py'):
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, f'"{sys.argv[0]}"', None, 1
                )
            else:
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.argv[0], "", None, 1
                )
            self.root.quit()
        except Exception as e:
            messagebox.showerror("错误", f"无法请求管理员权限: {e}")
    
    def check_admin_before_action(self):
        """在执行需要管理员权限的操作前检查"""
        try:
            import ctypes
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if not is_admin:
                response = messagebox.askyesno(
                    "需要管理员权限",
                    "此操作需要管理员权限！\n\n是否重新以管理员身份启动程序？"
                )
                if response:
                    self.restart_as_admin()
                return False
            return True
        except:
            return True
    
    def create_widgets(self):
        # 主标题
        title_label = ctk.CTkLabel(
            self.root, 
            text="Windows 端口管理工具", 
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title_label.pack(pady=10)
        
        # 添加广告信息（放在标题下方）
        self.create_ad_section()
        
        # 创建主框架
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=20, pady=(5, 10))
        
        # 创建选项卡
        self.create_tabs(main_frame)
        
    def create_ad_section(self):
        """创建广告区域"""
        ad_frame = ctk.CTkFrame(self.root, fg_color="#2b2b2b", border_width=1, border_color="orange")
        ad_frame.pack(fill="x", padx=20, pady=5)
        
        # 第一行：服务器广告
        server_frame = ctk.CTkFrame(ad_frame, fg_color="transparent")
        server_frame.pack(fill="x", padx=5, pady=3)
        
        server_label = ctk.CTkLabel(
            server_frame, 
            text="🚀 大宽带独享免备案免sm云服务器:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="orange"
        )
        server_label.pack(side="left", padx=5, pady=2)
        
        server_link = ctk.CTkButton(
            server_frame,
            text="https://xunduyun.com/",
            command=self.open_server_website,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#1f6aa5",
            hover_color="#0f4c75",
            text_color="white",
            width=180,
            height=25
        )
        server_link.pack(side="left", padx=3, pady=2)
        
        # 第二行：QQ群广告
        qq_label = ctk.CTkLabel(
            ad_frame,
            text="💬 QQ技术交流群: 262430517",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#00d4aa"
        )
        qq_label.pack(pady=(0, 5))
        
    def open_server_website(self):
        """打开服务器网站"""
        import webbrowser
        webbrowser.open("https://xunduyun.com/")
        
    def create_tabs(self, parent):
        # 创建选项卡视图
        self.tabview = ctk.CTkTabview(parent, width=750, height=500)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 添加选项卡
        self.tabview.add("端口管理")
        self.tabview.add("防火墙控制")
        self.tabview.add("系统信息")
        
        # 设置默认选项卡
        self.tabview.set("端口管理")
        
        # 创建各个选项卡的内容
        self.create_port_tab()
        self.create_firewall_tab()
        self.create_info_tab()
        
    def create_port_tab(self):
        """创建端口管理选项卡"""
        port_frame = self.tabview.tab("端口管理")
        
        # 开放单个端口部分
        single_port_frame = ctk.CTkFrame(port_frame)
        single_port_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(single_port_frame, text="开放单个端口", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        
        # 端口输入区域
        input_frame = ctk.CTkFrame(single_port_frame)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(input_frame, text="端口号:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.port_entry = ctk.CTkEntry(input_frame, placeholder_text="例如: 8080")
        self.port_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(input_frame, text="协议:").grid(row=0, column=2, padx=10, pady=5, sticky="w")
        self.protocol_var = ctk.StringVar(value="TCP")
        protocol_menu = ctk.CTkOptionMenu(input_frame, values=["TCP", "UDP", "Both"], variable=self.protocol_var)
        protocol_menu.grid(row=0, column=3, padx=10, pady=5, sticky="ew")
        
        input_frame.grid_columnconfigure(1, weight=1)
        
        # 按钮区域
        button_frame = ctk.CTkFrame(single_port_frame)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        open_port_btn = ctk.CTkButton(
            button_frame, 
            text="开放端口", 
            command=self.open_single_port,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        open_port_btn.pack(side="left", padx=10, pady=10)
        
        close_port_btn = ctk.CTkButton(
            button_frame, 
            text="关闭端口", 
            command=self.close_single_port,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="red",
            hover_color="darkred"
        )
        close_port_btn.pack(side="left", padx=10, pady=10)
        
        # 批量操作部分
        batch_frame = ctk.CTkFrame(port_frame)
        batch_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(batch_frame, text="批量操作", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        
        batch_button_frame = ctk.CTkFrame(batch_frame)
        batch_button_frame.pack(fill="x", padx=10, pady=10)
        
        open_all_btn = ctk.CTkButton(
            batch_button_frame,
            text="开放所有端口",
            command=self.open_all_ports,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="green",
            hover_color="darkgreen"
        )
        open_all_btn.pack(side="left", padx=10, pady=10)
        
        close_all_btn = ctk.CTkButton(
            batch_button_frame,
            text="取消开放所有端口",
            command=self.close_all_ports,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="orange",
            hover_color="darkorange"
        )
        close_all_btn.pack(side="left", padx=10, pady=10)
        
        # 状态显示区域
        self.status_text = ctk.CTkTextbox(port_frame, height=350)
        self.status_text.pack(fill="both", expand=True, padx=10, pady=10)
        self.status_text.insert("0.0", "欢迎使用端口管理工具！\n请以管理员身份运行以获得完整功能。\n")
        
    def create_firewall_tab(self):
        """创建防火墙控制选项卡"""
        firewall_frame = self.tabview.tab("防火墙控制")
        
        # 防火墙控制标题
        ctk.CTkLabel(firewall_frame, text="Windows防火墙控制", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        
        # 警告信息
        warning_frame = ctk.CTkFrame(firewall_frame)
        warning_frame.pack(fill="x", padx=10, pady=10)
        
        warning_label = ctk.CTkLabel(
            warning_frame, 
            text="⚠️ 警告：关闭防火墙可能会使您的计算机面临安全风险！",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="orange"
        )
        warning_label.pack(pady=10)
        
        # 防火墙状态
        status_frame = ctk.CTkFrame(firewall_frame)
        status_frame.pack(fill="x", padx=10, pady=10)
        
        self.firewall_status_label = ctk.CTkLabel(status_frame, text="防火墙状态: 检查中...", font=ctk.CTkFont(size=14))
        self.firewall_status_label.pack(pady=10)
        
        # 控制按钮
        control_frame = ctk.CTkFrame(firewall_frame)
        control_frame.pack(fill="x", padx=10, pady=10)
        
        disable_firewall_btn = ctk.CTkButton(
            control_frame,
            text="关闭防火墙",
            command=self.disable_firewall,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="red",
            hover_color="darkred"
        )
        disable_firewall_btn.pack(side="left", padx=10, pady=10)
        
        enable_firewall_btn = ctk.CTkButton(
            control_frame,
            text="启用防火墙",
            command=self.enable_firewall,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="green",
            hover_color="darkgreen"
        )
        enable_firewall_btn.pack(side="left", padx=10, pady=10)
        
        check_status_btn = ctk.CTkButton(
            control_frame,
            text="检查状态",
            command=self.check_firewall_status,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        check_status_btn.pack(side="left", padx=10, pady=10)
        
        # 初始检查防火墙状态
        self.check_firewall_status()
        
    def create_info_tab(self):
        """创建系统信息选项卡"""
        info_frame = self.tabview.tab("系统信息")
        
        ctk.CTkLabel(info_frame, text="系统网络信息", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(10, 5))
        
        # 信息显示区域 - 不设置固定高度，让它自动适应
        self.info_text = ctk.CTkTextbox(info_frame)
        self.info_text.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        
        # 刷新按钮 - 确保有足够空间显示
        refresh_btn = ctk.CTkButton(
            info_frame,
            text="刷新信息",
            command=self.refresh_system_info,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        refresh_btn.pack(pady=(0, 15))
        
        # 初始加载系统信息
        self.refresh_system_info()
        
    def run_command(self, command, show_output=True):
        """运行系统命令"""
        try:
            # 使用适合Windows的编码
            result = subprocess.run(
                command, 
                shell=True, 
                capture_output=True, 
                text=True,
                encoding='gbk',  # Windows中文系统默认编码
                errors='replace'  # 替换无法解码的字符
            )
            # 安全处理可能为None的输出
            stdout = result.stdout or ""
            stderr = result.stderr or ""
            output = stdout + stderr
            
            if show_output:
                self.update_status(f"执行命令: {command}\n输出: {output}\n")
            return result.returncode == 0, output
        except Exception as e:
            error_msg = f"执行命令失败: {str(e)}\n"
            if show_output:
                self.update_status(error_msg)
            return False, error_msg
    
    def update_status(self, message):
        """更新状态显示"""
        def update_ui():
            try:
                self.status_text.insert("end", message + "\n")
                self.status_text.see("end")
            except:
                pass
        
        # 确保在主线程中更新UI
        try:
            self.root.after(0, update_ui)
        except:
            # 如果主线程不可用，直接更新（可能在主线程中调用）
            update_ui()
        
    def open_single_port(self):
        """开放单个端口"""
        if not self.check_admin_before_action():
            return
            
        port = self.port_entry.get().strip()
        protocol = self.protocol_var.get()
        
        if not port:
            messagebox.showerror("错误", "请输入端口号！")
            return
            
        try:
            port_num = int(port)
            if port_num < 1 or port_num > 65535:
                messagebox.showerror("错误", "端口号必须在1-65535之间！")
                return
        except ValueError:
            messagebox.showerror("错误", "请输入有效的端口号！")
            return
        
        def open_port_thread():
            protocols = ["TCP", "UDP"] if protocol == "Both" else [protocol]
            
            for proto in protocols:
                # 添加入站规则 - 使用更兼容的命令格式
                command_in = f'netsh advfirewall firewall add rule name="Allow_Port_{port}_{proto}_IN" protocol={proto} dir=in localport={port} action=allow'
                success_in, output_in = self.run_command(command_in)
                
                # 添加出站规则
                command_out = f'netsh advfirewall firewall add rule name="Allow_Port_{port}_{proto}_OUT" protocol={proto} dir=out localport={port} action=allow'
                success_out, output_out = self.run_command(command_out)
                
                if success_in and success_out:
                    self.update_status(f"✅ 成功开放端口 {port} ({proto})")
                else:
                    self.update_status(f"❌ 开放端口 {port} ({proto}) 失败")
                    if output_in:
                        self.update_status(f"入站规则错误: {output_in}")
                    if output_out:
                        self.update_status(f"出站规则错误: {output_out}")
        
        threading.Thread(target=open_port_thread, daemon=True).start()
        
    def close_single_port(self):
        """关闭单个端口"""
        if not self.check_admin_before_action():
            return
            
        port = self.port_entry.get().strip()
        protocol = self.protocol_var.get()
        
        if not port:
            messagebox.showerror("错误", "请输入端口号！")
            return
        
        def close_port_thread():
            protocols = ["TCP", "UDP"] if protocol == "Both" else [protocol]
            
            for proto in protocols:
                # 删除入站规则 - 尝试两种命名格式
                command_in1 = f'netsh advfirewall firewall delete rule name="Allow_Port_{port}_{proto}_IN"'
                command_in2 = f'netsh advfirewall firewall delete rule name="Port_{port}_{proto}_IN"'
                success_in1, _ = self.run_command(command_in1)
                success_in2, _ = self.run_command(command_in2)
                
                # 删除出站规则
                command_out1 = f'netsh advfirewall firewall delete rule name="Allow_Port_{port}_{proto}_OUT"'
                command_out2 = f'netsh advfirewall firewall delete rule name="Port_{port}_{proto}_OUT"'
                success_out1, _ = self.run_command(command_out1)
                success_out2, _ = self.run_command(command_out2)
                
                if success_in1 or success_in2 or success_out1 or success_out2:
                    self.update_status(f"✅ 成功关闭端口 {port} ({proto})")
                else:
                    self.update_status(f"❌ 关闭端口 {port} ({proto}) 失败或规则不存在")
        
        threading.Thread(target=close_port_thread, daemon=True).start()
        
    def open_all_ports(self):
        """开放所有端口"""
        if not self.check_admin_before_action():
            return
            
        result = messagebox.askyesno("确认", "确定要开放所有端口吗？这可能会带来安全风险！")
        if not result:
            return
            
        def open_all_thread():
            self.update_status("🔥 开始开放所有端口...")
            
            # 创建允许所有入站连接的规则
            commands = [
                'netsh advfirewall firewall add rule name="Allow_All_TCP_IN" dir=in action=allow protocol=TCP',
                'netsh advfirewall firewall add rule name="Allow_All_UDP_IN" dir=in action=allow protocol=UDP',
                'netsh advfirewall firewall add rule name="Allow_All_TCP_OUT" dir=out action=allow protocol=TCP',
                'netsh advfirewall firewall add rule name="Allow_All_UDP_OUT" dir=out action=allow protocol=UDP'
            ]
            
            for command in commands:
                success, output = self.run_command(command)
                if success:
                    self.update_status(f"✅ 成功执行: {command.split('=')[-1]}")
                else:
                    self.update_status(f"❌ 执行失败: {command.split('=')[-1]}")
            
            self.update_status("🔥 所有端口开放完成！")
        
        threading.Thread(target=open_all_thread, daemon=True).start()
        
    def close_all_ports(self):
        """取消开放所有端口"""
        if not self.check_admin_before_action():
            return
            
        result = messagebox.askyesno("确认", "确定要取消开放所有端口吗？\n这将删除之前创建的所有端口开放规则。")
        if not result:
            return
            
        def close_all_thread():
            self.update_status("🔥 开始取消开放所有端口...")
            
            # 删除批量开放时创建的规则
            commands = [
                'netsh advfirewall firewall delete rule name="Allow_All_TCP_IN"',
                'netsh advfirewall firewall delete rule name="Allow_All_UDP_IN"',
                'netsh advfirewall firewall delete rule name="Allow_All_TCP_OUT"',
                'netsh advfirewall firewall delete rule name="Allow_All_UDP_OUT"'
            ]
            
            success_count = 0
            for command in commands:
                success, output = self.run_command(command)
                if success:
                    success_count += 1
                    self.update_status(f"✅ 成功删除规则: {command.split('=')[-1].replace('"', '')}")
                else:
                    self.update_status(f"❌ 删除规则失败或不存在: {command.split('=')[-1].replace('"', '')}")
            
            if success_count > 0:
                self.update_status(f"🔥 取消开放所有端口完成！成功删除 {success_count} 个规则")
            else:
                self.update_status("⚠️ 未找到需要删除的批量端口规则")
        
        threading.Thread(target=close_all_thread, daemon=True).start()
        
    def disable_firewall(self):
        """关闭防火墙"""
        if not self.check_admin_before_action():
            return
            
        result = messagebox.askyesno("确认", "确定要关闭Windows防火墙吗？这会降低系统安全性！")
        if not result:
            return
            
        def disable_firewall_thread():
            self.update_status("🔥 正在关闭防火墙...")
            
            commands = [
                'netsh advfirewall set allprofiles state off',
                'netsh advfirewall set domainprofile state off',
                'netsh advfirewall set privateprofile state off',
                'netsh advfirewall set publicprofile state off'
            ]
            
            for command in commands:
                success, output = self.run_command(command)
                if success:
                    self.update_status("✅ 防火墙已关闭")
                    break
            else:
                self.update_status("❌ 关闭防火墙失败")
            
            self.check_firewall_status()
        
        threading.Thread(target=disable_firewall_thread, daemon=True).start()
        
    def enable_firewall(self):
        """启用防火墙"""
        if not self.check_admin_before_action():
            return
            
        def enable_firewall_thread():
            self.update_status("🔥 正在启用防火墙...")
            
            command = 'netsh advfirewall set allprofiles state on'
            success, output = self.run_command(command)
            
            if success:
                self.update_status("✅ 防火墙已启用")
            else:
                self.update_status("❌ 启用防火墙失败")
            
            self.check_firewall_status()
        
        threading.Thread(target=enable_firewall_thread, daemon=True).start()
        
    def check_firewall_status(self):
        """检查防火墙状态"""
        def check_status_thread():
            # 使用更简单的命令检查防火墙状态
            command = 'netsh advfirewall show currentprofile state'
            success, output = self.run_command(command, show_output=False)
            
            if success and output:
                # 解析输出 - 检查中文和英文状态
                output_upper = output.upper()
                if "开" in output or "ON" in output_upper or "启用" in output:
                    status = "🟢 已启用"
                    color = "green"
                elif "关" in output or "OFF" in output_upper or "关闭" in output:
                    status = "🔴 已关闭" 
                    color = "red"
                else:
                    status = "❓ 状态未知"
                    color = "orange"
            else:
                # 备用检查方法
                command2 = 'netsh firewall show state'
                success2, output2 = self.run_command(command2, show_output=False)
                if success2 and output2:
                    if "启用" in output2 or "ON" in output2.upper():
                        status = "🟢 已启用"
                        color = "green"
                    else:
                        status = "🔴 已关闭"
                        color = "red"
                else:
                    status = "❓ 无法检查"
                    color = "orange"
            
            # 更新UI（需要在主线程中执行）
            def update_firewall_status():
                try:
                    self.firewall_status_label.configure(
                        text=f"防火墙状态: {status}",
                        text_color=color
                    )
                except:
                    pass
            
            try:
                self.root.after(0, update_firewall_status)
            except:
                update_firewall_status()
        
        threading.Thread(target=check_status_thread, daemon=True).start()
        
    def refresh_system_info(self):
        """刷新系统信息"""
        def refresh_info_thread():
            info_text = "正在获取系统网络信息...\n\n"
            
            def update_loading():
                try:
                    self.info_text.delete("0.0", "end")
                    self.info_text.insert("0.0", info_text)
                except:
                    pass
            
            try:
                self.root.after(0, update_loading)
            except:
                update_loading()
            
            # 获取网络接口信息
            success, output = self.run_command('ipconfig /all', show_output=False)
            if success:
                info_text += "=== 网络配置信息 ===\n" + output + "\n\n"
            
            # 获取端口监听信息
            success, output = self.run_command('netstat -an', show_output=False)
            if success:
                info_text += "=== 端口监听状态 ===\n" + output[:2000] + "...\n\n"
            
            # 获取防火墙规则
            success, output = self.run_command('netsh advfirewall firewall show rule name=all', show_output=False)
            if success:
                info_text += "=== 防火墙规则 ===\n" + output[:1000] + "...\n"
            
            # 更新显示
            def update_info_display():
                try:
                    self.info_text.delete("0.0", "end")
                    self.info_text.insert("0.0", info_text)
                except:
                    pass
            
            try:
                self.root.after(0, update_info_display)
            except:
                update_info_display()
        
        threading.Thread(target=refresh_info_thread, daemon=True).start()
        
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

if __name__ == "__main__":
    app = PortManagerApp()
    app.run() 