import cv2
import win32gui
import win32ui
import win32con
import numpy


def find_window(title: str) -> int:
    """
    获取窗口句柄
    :param title:
    :return:
    """
    hwnd = win32gui.FindWindow(None, title)
    if not win32gui.IsWindow(hwnd):
        raise Exception('句柄无效')
    return hwnd


def set_window_size(hwnd: int, width: int, height: int):
    """
    设置窗口大小
    :param rate:
    :param hwnd:
    :param width:
    :param height:
    :return:
    """
    win32gui.SetWindowPos(hwnd, None, 0, 0, width, height, 0)


def get_window_image(hwnd: int, width: int, height: int):
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    # 创建位图对象
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc_dc, width, height)

    save_dc.SelectObject(bmp)
    save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)

    # bmp.SaveBitmapFile(save_dc, 'D:/pythonProject/src/123.bmp')

    # 转换为 NumPy 数组
    bmp_info = bmp.GetInfo()
    bmp_str = bmp.GetBitmapBits(True)
    img = numpy.frombuffer(bmp_str, dtype=numpy.uint8)
    img.shape = (bmp_info['bmHeight'], bmp_info['bmWidthBytes'] // 4, 4)

    # BGR 转换为 RGB
    img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

    cv2.imwrite('src/window_image.png', img)

    # 清理资源
    win32gui.DeleteObject(bmp.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)
    return img


def find_image_in_window(hwnd, template_path, width: int, height: int):
    # 读取模板图像
    template = cv2.imread(template_path)
    if template is None:
        print("无法读取模板图像")
        return None

    # print(template)

    # 获取窗口图像
    window_image = get_window_image(hwnd, width, height)

    # 执行模板匹配
    result = cv2.matchTemplate(window_image, template, cv2.TM_CCOEFF_NORMED)
    threshold = 0.5  # 根据需要调整阈值
    y_loc, x_loc = numpy.where(result >= threshold)

    # 返回匹配位置
    if len(x_loc) > 0:
        return [(x, y) for (x, y) in zip(x_loc, y_loc)]
    else:
        return None


def save_matched_image(hwnd, template_path, width: int, height: int):
    positions = find_image_in_window(hwnd, template_path, width, height)
    if positions:
        # 获取窗口图像
        window_image = get_window_image(hwnd, width, height)

        for pos in positions:
            x, y = pos
            h, w = cv2.imread(template_path).shape[:2]
            # 裁剪出匹配的区域
            matched_image = window_image[y:y + h, x:x + w]
            cv2.imwrite('matched_image.png', matched_image)  # 保存匹配的图像
    else:
        print("未找到匹配的图像.")


def get_window_child(hwnd: int):
    if not hwnd:
        raise Exception('传入句柄不存在')
    hwnd_child_list = []
    win32gui.EnumChildWindows(hwnd, lambda item, param: param.append(item), hwnd_child_list)
    return hwnd_child_list


