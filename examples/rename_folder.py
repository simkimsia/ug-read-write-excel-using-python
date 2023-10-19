import os
import glob

current_dir = os.path.dirname(os.path.realpath(__file__))

# for fn in os.listdir(current_dir):
#     if not os.path.isdir(os.path.join(current_dir, fn)):
#         # Not a directory
#         continue
#     if fn.startswith('c9'):
#         parts = fn.partition('_')
#         # this should cxx
#         prefix = parts[0]
#         underscore = parts[1]
#         suffix = parts[2]
#         new_name = 'c0' + prefix[1] + \
#             underscore + prefix[2:] + underscore + suffix
#         os.rename(os.path.join(current_dir, fn),
#                   os.path.join(current_dir, new_name))

test_path = os.path.join(os.path.dirname(current_dir), 'tests')
test_path_py = os.path.join(test_path, '*.py')

for filename in glob.iglob(test_path_py, recursive=True):
    print(filename)
    basename = os.path.basename(filename)
    if basename.startswith('test_9'):
        new_name = f'test_09{basename[6:]}'
        os.rename(filename,
                  os.path.join(test_path, new_name))


# for fn in os.listdir(current_dir):
#     if not os.path.isdir(os.path.join(current_dir, fn)):
#         # Not a directory
#         continue
#     if fn.startswith('c0'):
#         for filename in glob.iglob(test_path_py, recursive=True):
#             with open(filename, 'r+') as f:
#                 for line in f:
#                     searchFor = 'c07_' + fn[4:]
#                     if searchFor in line:
#                         f.write(line.replace(searchFor, fn))
#                     searchFor = 'c08_' + fn[4:]
#                     if searchFor in line:
#                         f.write(line.replace(searchFor, fn))
#                     searchFor = 'c9' + fn[4:]
#                     if searchFor in line:
#                         f.write(line.replace(searchFor, fn))
#             f.close()
