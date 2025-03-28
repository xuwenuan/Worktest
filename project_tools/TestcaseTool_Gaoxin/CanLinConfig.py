import ctypes
import struct

class CanLinConfig:
    '''
       startBit 配置项startBit， type： int
       length 配置项length， type： int
       value 配置项值， type int (其他为10进制值)
       dlc msg长度
       '''
    def getMessage(self, startBit, length, value, dlc):
        startBit = int(startBit)
        length = int(length)
        value = int(value)
        byte_array = bytearray(int(dlc))

        return self.calc_new_value(byte_array, startBit, length, value, encoding_order="motorola")

    '''
    startByte 配置项startByte， type： int
    startBit 配置项startBit， type： int
    length 配置项length， type： int
    value 配置项值， type string (诊断ID为16进制值， 其他为10进制值)
    dlc msg长度
    '''
    def getConfigBytesString(self, startByte, startBit, length, value, dlc=11):
        byte_array = bytearray(dlc)
        for i in range(dlc):
            byte_array[i] = 0xFF

        if length < 8:
            value = self.bits_complement(int(value), 2)
            value = value << (startBit % 8)
            byte_array[startByte] = self.bits_complement(value, 8)
        elif length == 8:
            byte_array[startByte] = int(value)
        elif length > 8:
            value = value.upper().replace("0X", "").zfill(4)
            value1 = value[:-2]
            value2 = value[-2:]
            byte_array[startByte] = int(value1, 16)
            byte_array[startByte + 1] = int(value2, 16)

        byteString = ''
        for item in byte_array:
            byteString = byteString + " " + hex(item).upper().replace("0X", "").zfill(2)
        return byteString.strip()

    def string_to_uint(self, s):
        return ctypes.c_uint(int(s)).value

    def float_to_hex(self, float_num):
        # 将浮点数转换为4字节序列
        bytes_obj = struct.pack('>f', float_num)
        # 将字节序列转换为十六进制字符串
        hex_str = bytes.hex(bytes_obj)
        return hex_str

    def getSignalLenAndSignalValue(self, signalType, inputValue):
        signalLen = " 00 00 00"
        tt1 = '1'
        tt2 = '0'
        v = 1
        if 'int32' in signalType.lower():
            v = 4
            tt1 = '3'
        elif 'float' in signalType.lower():
            v = 4
            tt1 = '4'
        elif 'int16' in signalType.lower():
            v = 2
            tt1 = '2'
        if signalType.lower().startswith('int') or signalType.lower().startswith('sint') or signalType.lower().startswith('float'):
            tt2 = '8'
        tt = tt2 + tt1
        signalLen = " " + tt + signalLen #+ str(v)
        try:
        # if 1:
            tt3 = '0'
            if 'uint' in signalType.lower():
                signalValue = hex(self.string_to_uint(inputValue.strip())).lower().replace("0x", "").zfill(8)
            elif 'float' in signalType.lower():
                tempValue = float(inputValue.strip())
                if tempValue < 0:
                    tt3 = '8'
                signalValue = self.float_to_hex(tempValue)
            else:
                tempValue = int(inputValue.strip())
                if tempValue >= 0:
                    signalValue = hex(tempValue).lower().replace("0x", "").zfill(8)
                else:
                    tt3 = '8'
                    tempValue = 0 - tempValue
                    # tempValue = (1 << (v - 1) * 8 + 7) | tempValue
                    signalValue = hex(tempValue).lower().replace("0x", "").zfill(8)

            signalValue = ' ' + tt3 + tt1 + ' ' + ' '.join(signalValue[i:i + 2] for i in range(0, len(signalValue), 2))
        except:
            signalLen = " 00 00 00 00"
            signalValue = " 00 00 00 00"
        return signalLen, signalValue

    def bits_complement(self, number, bits):
        # print(number,bits,~number & ((1 << bits) - 1))
        return ~number & ((1 << bits) - 1)

    def getStartandLengthHex(self, start, length):
        hex_str = hex(int(start)).upper().replace("0X", "").zfill(4) + hex(int(length)).upper().replace("0X", "").zfill(4)
        return ' ' + ' '.join(hex_str[i:i + 2] for i in range(0, len(hex_str), 2))

    def calc_new_value(self, data, start_bit, bit_len, target_value, encoding_order="intel"):
        """
        根据格式，修改指定位的数据
        :param data: 待修改的原始数据
        :param start_bit: 起始bit位
        :param bit_len: 占位长度
        :param target_value: 目标值
        :param encoding_order: motorola、intel
        """

        def padding_zero(value, n=8):
            """
            补零操作
            :param value: 需要补零的数据
            :param n: 需要补齐的位数，默认补齐8位
            :return: 补零后的数据
            """
            return value if len(value) == n else '0' * (n - len(value)) + value



        def get_byte_list(start_bit, bit_len, encoding_order="intel"):
            """
            根据起始位和长度，计算出当前需要修改哪些字节
            :param start_bit: 起始位
            :param bit_len: 长度
            :return: 字节索引列表
            """
            # 计算需要修改的byte  -s
            byte_set = set()
            for i in range(bit_len):
                byte_set.add(int((start_bit + i) / 8))
            byte_list = sorted(list(byte_set))

            if encoding_order != 'intel':
                # motorola格式，索引号是从前往后的，所以使用加减法的方式进行换算
                # intel下的[2，3]，对应到motorola下的[1, 2]
                # intel下的[2，3, 4]，对应到motorola下的[0, 1, 2]
                byte_list = [i - (len(byte_list) - 1) for i in byte_list]
            # 计算需要修改的byte  -e
            return byte_list

        def byte_join(byte_list, data, encoding_order="intel"):
            """
            将data中的数据根据byte_list进行拼接
            :param byte_list: 待处理的byte
            :param data: 所有数据
            :return: 拼接后的二进制数据
            """
            final_str = ''
            # intel:将后面的byte放到前面的byte的前面
            # motorola:将前面的byte放到后面的byte的前面
            if encoding_order == "intel":
                byte_list = byte_list[::-1]
            for i in range(len(byte_list)):
                byte_num = byte_list[i]
                byte_value = data[byte_num]
                byte_value_bin = bin(byte_value)[2:]
                byte_value_bin = padding_zero(byte_value_bin)
                final_str += byte_value_bin
            final_list = list(final_str)
            return final_list

        byte_list = get_byte_list(start_bit, bit_len, encoding_order)
        #print("需要修改的byte=", byte_list)

        final_list = byte_join(byte_list, data, encoding_order)
        #print("拼接后的二进制列表=", final_list)

        # 将拼接后的数据进行倒置，与列表形式对齐
        final_list_reverse = final_list[::-1]
        #print("翻转后的二进制列表=", final_list_reverse)

        target_value_bin = list(bin(target_value)[2:])
        target_value_bin_reverse = target_value_bin[::-1]
        # print('翻转后的期望修改的值=', target_value_bin_reverse)

        # 根据start_bit向零对齐处理，找到开始进行修改的位:第一个byte对应的bit索引
        if encoding_order == 'intel':
            # intel模式下，用第一个byte拿到索引
            edit_bit = start_bit - 8 * byte_list[0]
        else:
            # motorola模式下，用最后一个byte拿到索引
            edit_bit = start_bit - 8 * byte_list[-1]
        # 遍历需要修改的长度进行取值
        for i in range(bit_len):

            final_list_reverse[edit_bit + i] = target_value_bin_reverse[i] if i < len(target_value_bin_reverse) else '0'
        final_list = final_list_reverse[::-1]
        #print("修改后的二进制列表=", final_list)

        # 需要将byte_list倒序，因为叠加的时候是后面的叠加到了前面
        if encoding_order == "intel":
            byte_list = byte_list[::-1]
        for i in range(len(byte_list)):
            split_value = ''.join(final_list[i + 7 * i:i + 7 * i + 8])
            data[byte_list[i]] = int(split_value, 2)
        #print("修改后的完整数据=", data)
        temparray = []
        for d in data:
            temparray.append(hex(d).upper().replace("0X", "").zfill(2))
        return " ".join(temparray).strip()




if __name__ == '__main__':
    config = CanLinConfig()
    signalLen, signalValue = config.getSignalLenAndSignalValue('int16', '-1')
    print(signalLen)
    print(signalValue)