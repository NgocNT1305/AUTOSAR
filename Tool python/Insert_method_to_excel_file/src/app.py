import os
import shutil
import glob

 

def get_current_directory():

    """Trả về thư mục hiện tại của file main.py."""

    return os.path.dirname(os.path.abspath(__file__))


def delete_files_in_directory(directory):

    """Xóa hết các file trong thư mục được chỉ định."""

    for filename in os.listdir(directory):

        file_path = os.path.join(directory, filename)

        if os.path.isfile(file_path):

            os.remove(file_path)


def copy_files(src_directory, dest_directory):

    """Sao chép các file Excel từ thư mục nguồn sang thư mục đích và thêm đuôi _copy vào tên file."""

    for file_path in glob.glob(os.path.join(src_directory, '*.xlsx')):

        file_name = os.path.basename(file_path)

        dest_file_name = file_name.replace('.xlsx', '_copy.xlsx')

        dest_file_path = os.path.join(dest_directory, dest_file_name)

        shutil.copy(file_path, dest_file_path)

    return file_name

 

def app_main(file_name):

    # Chuyển đổi các đường dẫn thành đường dẫn tuyệt đối

    input_dir = os.path.abspath(os.path.join(get_current_directory(), '..', 'input'))

    output_dir = os.path.abspath(os.path.join(get_current_directory(), '..', 'output'))

    #Xoa toan bo file thuoc folder output

    delete_files_in_directory(output_dir)

    file_name = copy_files(input_dir,output_dir)

    return output_dir