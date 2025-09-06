import uiautomation as auto
import time
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def find_wechat_window():
    """查找微信窗口，支持不同标题"""
    possible_titles = ["微信", "微信 WeChat"]
    for title in possible_titles:
        window = auto.WindowControl(Name=title)
        if window.Exists(2):
            logging.info(f"找到微信窗口: {title}")
            return window
    # 尝试通过类名查找
    window = auto.WindowControl(ClassName="WeChatMainWndForPC")
    if window.Exists(2):
        logging.info(f"通过类名找到微信窗口")
        return window
    return None


def send_to_wechat_assistant(message):
    try:
        # 获取微信窗口
        wechat_window = find_wechat_window()
        if not wechat_window:
            logging.error("未找到微信窗口，请确保微信已启动")
            return False

        # 查找搜索框（多种可能的名称）
        search_box = None
        for name in ["搜索", "搜索联系人、公众号"]:
            search_box = wechat_window.EditControl(Name=name)
            if search_box.Exists(1):
                logging.info(f"找到搜索框: {name}")
                break

        if not search_box or not search_box.Exists():
            logging.error("未找到搜索框")
            return False

        # 点击搜索框并输入
        search_box.Click(simulateMove=False)
        search_box.SendKeys("{Ctrl}a{Delete}")  # 清空搜索框
        time.sleep(0.5)
        search_box.SendKeys("jyr", waitTime=0.2)
        time.sleep(1.5)  # 增加等待时间确保搜索结果加载

        # 查找搜索结果中的联系人项
        assistant_item = None
        # 使用更精确的控件类型和名称匹配
        for ctrl in wechat_window.ListControl().GetChildren():
            if isinstance(ctrl, auto.ListItemControl) and "jyr" in ctrl.Name:
                assistant_item = ctrl
                break

        if not assistant_item or not assistant_item.Exists(3):
            logging.error("未找到'jyr'联系人")
            return False

        # 打开聊天窗口 - 使用更可靠的方式
        # 先确保联系人项可见
        assistant_item.SetFocus()
        time.sleep(0.5)

        # 尝试双击打开
        try:
            assistant_item.DoubleClick(simulateMove=False)
        except:
            # 如果双击失败，尝试回车键
            assistant_item.SendKeys("{Enter}")

        time.sleep(1.5)  # 等待聊天窗口打开

        # 查找输入框 - 添加更多可能的名称
        input_box = None
        input_names = ["输入消息", "请输入", "发送消息"]
        for name in input_names:
            input_box = wechat_window.EditControl(Name=name)
            if input_box.Exists(1):
                logging.info(f"找到输入框: {name}")
                break

        if not input_box or not input_box.Exists():
            logging.error("未找到输入框")
            return False

        # 发送消息
        input_box.SendKeys(message, waitTime=0.1)
        time.sleep(0.5)

        # 尝试多种发送方式
        try:
            # 方式1: Ctrl+Enter
            input_box.SendKeys("{Ctrl}{Enter}", waitTime=0.2)
        except:
            try:
                # 方式2: 点击发送按钮
                send_btn = wechat_window.ButtonControl(Name="发送")
                if send_btn.Exists():
                    send_btn.Click()
                else:
                    # 方式3: Enter键
                    input_box.SendKeys("{Enter}", waitTime=0.2)
            except Exception as e:
                logging.warning(f"发送按键失败: {str(e)}")
                return False

        logging.info("消息发送成功")
        return True

    except Exception as e:
        logging.error(f"发送过程出错: {str(e)}", exc_info=True)
        return False


if __name__ == "__main__":
    # 确保以管理员权限运行
    try:
        import ctypes

        if not ctypes.windll.shell32.IsUserAnAdmin():
            logging.warning("建议以管理员权限运行程序以获得更好的兼容性")
    except:
        pass

    # 测试发送
    send_success = send_to_wechat_assistant("测试消息：Hello，这是一条自动发送的消息！")
    if not send_success:
        logging.info("尝试使用备用方法...")
        # 改进的备用方法
        auto.SendKeys("{Ctrl}f", waitTime=0.5)
        time.sleep(1)
        auto.SendKeys("jyr", waitTime=0.5)
        time.sleep(1.5)  # 确保搜索结果出现
        auto.SendKeys("{Down}")  # 选择第一个搜索结果
        time.sleep(0.5)
        auto.SendKeys("{Enter}")  # 进入聊天
        time.sleep(1)
        auto.SendKeys("备用方法发送的消息")
        time.sleep(0.5)
        auto.SendKeys("{Enter}")  # 发送消息