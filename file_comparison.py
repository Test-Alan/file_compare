import os

'''
    需求说明：
    某个目录下有若干个文件和目录，文件名和目录名是一样的
    需要知道文件名和目录名是否都一样并且数量是一样多的。
    实现
    1.获取某个目录下的所有文件名和目录名
    2.然后对比两个名称是否相同，数量是否一样
    3.把不一样的显示出来
'''
# 获取某个路径的所有文件名和目录名
def get_file_name(file_dir):

    for root, dirs, files in os.walk(file_dir):

        # print(root)  # 当前目录路径
        # print(dirs)  # 当前路径下所有子目录
        # print(files)  # 当前路径下所有非目录子文件
        return dirs , files


# 去除文件名的后缀
def get_remove_suffix_filename(file_path):
    # 获取当前目录下的所有文件名
    file_name = get_file_name(file_path)[1]

    # 因为文件名是带后缀的需要把后缀去掉，只取名称然后把文件名添加到新的列表中
    new_file_name = []
    for i in file_name:
        name = i.split('.')
        name = name[0]
        new_file_name.append(name)
    return new_file_name

# 对比当前目录下的文件名是否和目录名相同并且数量相等
def contrast_file(files_path, dirs_path):
    # 获取当前目录下的所有文件名
    files_name = get_remove_suffix_filename(files_path)
    # 获取当前目录下的所有目录名
    dirs_name = get_file_name(dirs_path)[0]
    print('表格名字', files_name )
    print('合同名字', dirs_name)
    print('表格有的合同没有的:', list(set(files_name ).difference(set(dirs_name))))
    print('合同有的表格没有是的:', list(set(dirs_name).difference(set(files_name ))))
def main():
    path1 = r'E:\代偿用户对比\对比文件\表格'  # 文件1
    path2 = r'E:\代偿用户对比\对比文件\合同'  # 文件2
    contrast_file(path1, path2)



if __name__ == '__main__':
    main()

