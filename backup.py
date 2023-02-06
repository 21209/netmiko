from multiprocessing.pool import ThreadPool
from openpyxl import load_workbook
from netmiko import Netmiko
import logging
import time
import os
import re

# 保存日志文件，设备备份失败时诊断错误用，在hosts.xlsx 中将可正常备份的设备注释掉，每次测试只保留一个（因为脚本登陆设备是并行执行的，如果保留多个，日志内容顺序会混乱无意义）
# logging.basicConfig(filename=f"{time.strftime('%Y-%m-%d-%H%M')}-run.log", level=logging.DEBUG)
# logger = logging.getLogger("netmiko")

# 备份文件保存目录，年份》月份》日期
year_dir = time.strftime('%Y')
month_dir = time.strftime('%m')
day_dir = time.strftime('%d')

root = os.getcwd() 

if not os.path.exists(year_dir):
    os.mkdir(year_dir)
    os.chdir(year_dir)
    if not os.path.exists(month_dir):
        os.mkdir(month_dir)
        os.chdir(month_dir)
        if not os.path.exists(day_dir):
            os.mkdir(day_dir)
            os.chdir(day_dir)

os.chdir(root)
       
root = os.path.join(root, year_dir, month_dir, day_dir)


index = [chr(i) for i in range(97, 123)]
index = index[:9]

failed = []

if not os.path.exists(root):
    os.mkdir(root)


def read_excel(host_excel):
    if not os.path.exists(host_excel):
        print(f'没有 {host_excel}文件！')
        exit()

    wb = load_workbook(host_excel)
    sheet1 = wb['Sheet1']

    i = 0

    total = []
    list_host = []

    host_sum = 0
    for cell in sheet1['A']:
        if i == 0:
            i = 1
            continue
        elif cell.value is None or str(cell.value)[0] == '#':
            continue

        else:
            host_sum += 1
            part = []
            for i in index:
                part.append(sheet1[f'{i}{str(cell.coordinate)[1:]}'].value)
        total.append(part)

    print(f'一共{host_sum :<3}台设备{time.strftime("%Y-%m-%d"):>32}')
    print('-' * 52)

    for row in total:
        host_id = row[0]
        host_loc = row[1]
        host_type = row[3]
        host_ip = row[4]
        host_port = row[5]
        host_user = row[6]
        host_passwd = row[7]
        host_command = row[8]

        host = {
            "device_type": host_type,
            "host": host_ip,
            "port": host_port,
            "username": host_user,
            "password": host_passwd
        }
        pack = {
            "host_id": host_id,
            "host_loc": host_loc,
            "command": host_command
        }
        all_info = {
            "host": host,
            "pack": pack
        }
        list_host.append(all_info)
    return list_host, (host_sum - 1)


def login_backup(host_info):
    success = 0
    try:

        net_connect = Netmiko(**host_info['host'])
        prompt = net_connect.find_prompt()
        
        print(
            f"{host_info['pack']['host_id']:0>2} 号设备 {prompt:>23}登录成功 {host_info['host']['host']:>18} \n",
            end='')
        success += 1

    except Exception as e:
        print(e)
        print(f"{host_info['pack']['host_id']:0>2} 号设备 {host_info['host']['host']:>18} 登录失败 \n", end='')
        failed.append(f"{host_info['pack']['host_id']:0>2} 号设备 {host_info['host']['host']:>18} {host_info['pack']['host_loc']}")



    else:
        node = host_info['pack']['host_loc']
        save_dir = os.path.join(root, node)

        if not os.path.exists(save_dir):
            os.mkdir(save_dir)

        sys_name = net_connect.find_prompt().replace('<', '').replace('>', '')
        sys_name = sys_name.replace('#', '')

        filename = f"{host_info['pack']['host_id']}号-{sys_name}-[{host_info['host']['host']}]-{time.strftime('%Y-%m-%d-%H%M')}.cfg"

        output = net_connect.send_command_timing(host_info['pack']['command'])

        filename = os.path.join(save_dir, filename)

        # 当登录账户没有取消分页（screen-length 0 temporary）的权限时，需要向channel 中发送空格来保证接收到配置文件内容的完整性
        while True:
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)
            if "- More -" in output:
                output += net_connect.send_command_timing('        \n', strip_prompt=False, strip_command=False,
                                                          normalize=False)

            elif "- More -" not in output:
                break
            output = output.split("\n")
        # 目前锐捷设备有取消分页（screen-length 0 temporary）的权限，所以跳过发送空格的步骤
        if host_info['host']['device_type'] == 'ruijie_os':
            with open(file=filename, mode='a+', encoding='utf-8') as f2:
                f2.write(output)

        else:

            for line in output:
                # 删掉每行配置前的 --More-- 以及空格
                line = re.sub('\s*(-)+\sMore\s(-)+\s*', '', line)
                # H3C 部分型号读到的配置有ASCII 转义字符，给替换掉
                line = re.sub('^([^s]*)16D','',line)
                if prompt in line:
                    continue

                with open(file=filename, mode='a+', encoding='utf-8') as f2:
                    f2.write(line + '\n')

        net_connect.disconnect()
    return


if __name__ == "__main__":
    # 读取hosts.xlsx 文件内容
    host_info = read_excel('hosts.xlsx')
    login_info = host_info[0]
    host_nums = host_info[1]

    # 设置同一时间可连接交换机的最大数量
    pool_size = 30

    pool = ThreadPool(pool_size)
    pool.map(login_backup, login_info)
    pool.close()
    pool.join()

    print('-' * 52)
    print(f'{len(failed)} 台设备未备份成功:')
    print('-' * 52)
    with open(os.path.join(root, '备份失败设备.txt'), 'a+') as f:
        for device in failed:
            print(device)
            f.write(device)
            f.write('\r\n')
