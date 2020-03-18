import os
import zipfile


def file_to_zip(target_file, zip_file_name=''):
    """
    :param target_file:要压缩的文件
    :param zip_file_name: 压缩文件的名字
    :return:
    """
    fzip = zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED)
    if os.path.isfile(target_file):
        # 压缩单独文件
        fzip.write(zip_file_name)
    else:
        # 压缩文件夹
        for root, dirs, files in os.walk(target_file):
            froot = root.replace(target_file, '')  # 防止文件从根目录开始压缩，只保留给定文件夹后的路径
            for file_name in files:
                true_path = os.path.join(root, file_name)  # 要压缩的文件路径
                file_path = os.path.join(froot, file_name)  # 压缩文件中的文件路径
                fzip.write(true_path, file_path)
    fzip.close()


if __name__ == '__main__':
    file_to_zip("D:\\project\\project_exercises\\deal_excel", "D:\\project\\result.zip")
