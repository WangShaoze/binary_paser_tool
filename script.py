import ctypes
import base64


# 定义结构体（同之前，单个结构体大小1112字节）
class STExportValue_t(ctypes.Structure):
    _fields_ = [
        ("data_len", ctypes.c_int),
        ("data_buf", ctypes.c_char * 1024),  # 512*2=1024
        ("time_len", ctypes.c_int),
        ("time_buf", ctypes.c_char * 32),
        ("protocol_len", ctypes.c_int),
        ("protocol_type_buf", ctypes.c_char * 7),
        ("analysis_len", ctypes.c_int),
        ("analysis_result_buf", ctypes.c_char * 32),
    ]


# 单个结构体的字节大小（验证是否为1112）
STRUCT_SIZE = ctypes.sizeof(STExportValue_t)
assert STRUCT_SIZE == 1112, f"结构体大小错误，实际为{STRUCT_SIZE}字节"


def parse_multiple_structs(binary_data):
    """解析多组结构体数据，返回列表"""
    struct_list = []
    data_length = len(binary_data)

    # 检查总长度是否为结构体大小的整数倍（避免数据不完整）
    if data_length % STRUCT_SIZE != 0:
        raise ValueError(f"二进制数据长度不正确，应为{STRUCT_SIZE}的整数倍，实际为{data_length}")

    # 按结构体大小分割并解析
    for i in range(0, data_length, STRUCT_SIZE):
        # 截取当前结构体的二进制数据
        chunk = binary_data[i:i + STRUCT_SIZE]
        # 转换为结构体实例
        struct_instance = STExportValue_t.from_buffer_copy(chunk)
        struct_list.append(struct_instance)

    return struct_list


# ------------------------------
# 使用示例：解析包含多组数据的二进制文件
# ------------------------------
if __name__ == "__main__":
    # 读取二进制文件（假设包含多组结构体数据）
    with open("2025-10-28-10-04-09.txt", "rb") as f:
        all_binary_data = f.read()

    try:
        # 解析所有数据
        structs = parse_multiple_structs(all_binary_data)
        print(f"成功解析 {len(structs)} 组数据")

        # 遍历并打印第一组数据（示例）
        if structs:
            for idx, uni in enumerate(structs):
                print("====================== 第{}组数据: Start ============".format(idx + 1))
                print("data_len:", uni.data_len)
                data_buf_base64_str = uni.data_buf.decode('utf-8', errors='ignore')
                # 进行Base64解码（注意：Base64字符串需为字节类型，这里用utf-8编码转换）
                decoded_bytes = base64.b64decode(data_buf_base64_str.encode('utf-8'))
                # 将解码后的字节转换为字符串（假设是UTF-8编码）
                decoded_str = decoded_bytes.decode('utf-8')
                print("data_buf:", decoded_str)    # 对应 模板.xlsx 中的  源数据
                time_len = uni.time_len
                print("time_len:", time_len)
                time_buf = uni.time_buf.decode('utf-8', errors='ignore')    # 对应 模板.xlsx 中的  时间
                time_buf = " ".join([uni.strip() for uni in time_buf.strip().split("\n")])
                print("time_buf:", time_buf)
                protocol_len = uni.protocol_len
                print("protocol_len:", protocol_len)
                protocol_type_buf = uni.protocol_type_buf.decode('utf-8', errors='ignore')   # 对应 模板.xlsx 中的 协议类型
                print("protocol_type_buf:", protocol_type_buf)
                analysis_len = uni.analysis_len
                print("analysis_len:", analysis_len)
                analysis_result_buf = uni.analysis_result_buf.decode('utf-8', errors='ignore')  # 对应 模板.xlsx 中的 分析结果
                print("analysis_result_buf:", analysis_result_buf)
                print("====================== 第{}组数据: END ============".format(idx + 1))
    except ValueError as e:
        print(f"解析失败：{e}")
