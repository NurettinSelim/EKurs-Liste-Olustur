from course_list_generator import create_list, save_list_to_csv

list_name = input("List name:")
lists = create_list(f"excels/{list_name}.xlsx")

save_list_to_csv("csv/", lists)
