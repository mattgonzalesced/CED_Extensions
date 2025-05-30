# -*- coding: utf-8 -*-
import os.path

import clr
# pylint: disable=import-error,invalid-name,broad-except
from pyrevit import script

clr.AddReference('System')
from System.IO import Directory, File, Path

output = script.get_output()
logger = script.get_logger()

def copy_folder(source_dir, target_dir):
    output.print_md("🔄 **Copying from** `{}` **to** `{}`".format(source_dir, target_dir))

    # Create target directory if it doesn't exist
    if not Directory.Exists(target_dir):
        Directory.CreateDirectory(target_dir)
        output.print_md("- Created target directory `{}`".format(target_dir))

    # Copy all files
    for file_path in Directory.GetFiles(source_dir):
        file_name = Path.GetFileName(file_path)
        target_file = Path.Combine(target_dir, file_name)
        File.Copy(file_path, target_file, True)  # True = overwrite existing files
        output.print_md("- Copied file: `{}`".format(file_name))

    # Copy all subdirectories recursively
    for dir_path in Directory.GetDirectories(source_dir):
        dir_name = Path.GetFileName(dir_path)
        target_subdir = Path.Combine(target_dir, dir_name)
        output.print_md("- Entering subdirectory: `{}`".format(dir_name))
        copy_folder(dir_path, target_subdir)

def main():
    user_folder = os.path.expanduser('~')
    source_path = r"DC\ACCDocs\CoolSys\CED Content Collection\Project Files\Temp"
    target_path = r"OneDrive - CoolSys Inc\Desktop\_TARGET TEST"
    source_dir = os.path.join(user_folder, source_path)
    target_dir = os.path.join(user_folder, target_path)

    output.print_md("# 🚀 **Updating Extension from Source**")
    output.print_md("🔎 Source Path: `{}`".format(source_dir))
    output.print_md("📝 Target Path: `{}`".format(target_dir))

    # Step 1: Delete everything in target directory
    if Directory.Exists(target_dir):
        output.print_md("🗑️ **Deleting old content in target directory…**")
        # Delete all files
        for file_path in Directory.GetFiles(target_dir):
            File.Delete(file_path)
            output.print_md("- Deleted file: `{}`".format(Path.GetFileName(file_path)))

        # Delete all subdirectories
        for dir_path in Directory.GetDirectories(target_dir):
            Directory.Delete(dir_path, True)
            output.print_md("- Deleted directory: `{}`".format(Path.GetFileName(dir_path)))
    else:
        output.print_md("✅ Target directory did not exist. No deletions needed.")

    # Step 2: Copy everything from source to target
    output.print_md("📂 **Copying new content…**")
    copy_folder(source_dir, target_dir)

    output.print_md("🎉 **Update complete!**")

main()


# res = True
#
# if res:
#     logger = script.get_logger()
#     results = script.get_results()
#
#     # re-load pyrevit session.
#     logger.info('Reloading....')
#     sessionmgr.reload_pyrevit()
#
#     results.newsession = sessioninfo.get_session_uuid()
