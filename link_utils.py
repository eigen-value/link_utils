# ***** BEGIN GPL LICENSE BLOCK *****
#
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software Foundation,
# Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.
#
# ***** END GPL LICENCE BLOCK *****

bl_info = {
    "name": "Linked Library Utilities",
    "author": "Lucio Rossi",
    "version": (0, 0, 0),
    "blender": (2, 74, 0),
    "location": "View3D > Toolshelf > Linked Library Utilities",
    "description": "Lists objects linked from a .blend library.",
    "wiki_url": "",
    "category": "Object",
}

import bpy

# UI
#      Hide the entire panel for non-linked objects?


def get_linked_dict():
    linked_dict = {}
    linked_list = []

    for ob in bpy.context.scene.objects:
        if ob.dupli_group and ob.dupli_group.library:
            obpath = ob.dupli_group.library.filepath
            linked_list.append(ob.dupli_group.name)
            linked_dict[ob.dupli_group.name] = [obpath]
        if ob.library:
            obpath = ob.library.filepath
            linked_list.append(ob.name)
            linked_dict[ob.name] = [obpath]
        elif ob.proxy:
            ob = ob.proxy
            obpath = ob.library.filepath
            linked_list.append(ob.name)
            linked_dict[ob.name] = [obpath]

    for key in linked_dict.keys():
        linked_dict[key].append(linked_list.count(key))


    return linked_dict

def dict_to_txt(d,fname,headers):

        text = open(fname, 'w')
        for h in headers:
            text.write(h + ',')
        text.write('\n')

        keys = list(d.keys())
        keys.sort()

        for key in keys:
            text.write(key + ',')
            for i in range(len(d[key])):
                text.write(str(d[key][i]) + ',')
            text.write('\n')

def dict_to_csv(d,fname,headers):
    import csv

    with open(fname, 'w') as csvfile:
        writer = csv.writer(csvfile, dialect='excel')
        writer.writerow(headers)

        keys = list(d.keys())
        keys.sort()

        for key in keys:
            row = [key]
            row.extend(d[key])
            writer.writerow(row)

def dict_to_xlsx(d,fname,headers):

        import xlsxwriter

        xlsx_file = xlsxwriter.Workbook(fname)
        xlsx_worksheet = xlsx_file.add_worksheet()
        bold = xlsx_file.add_format({'bold': True})

        for i,h in enumerate(headers):
            xlsx_worksheet.write(0,i,h,bold)

        keys = list(d.keys())
        keys.sort()

        for i, key in enumerate(keys):
            xlsx_worksheet.write(i+1,0,key)
            for j in range(len(d[key])):
                xlsx_worksheet.write(i+1,j+1,d[key][j])

class SaveLinkedList(bpy.types.Operator):
    """Save Linked List"""
    bl_idname = "object.save_linked"
    bl_label = "Save Linked List"

    def execute(self, context):

        scene = context.scene
        path = bpy.path.abspath(scene.link_utils_path)
        fname = bpy.path.basename(bpy.context.blend_data.filepath)
        fname = fname[:-6] # cut off extension
        linked_dict = get_linked_dict()
        headers = ['Obj name','Path','num']



        if scene.link_utils_file_ext == '0':
            fname = fname + '_linked_list.txt'
            dict_to_txt(linked_dict,path + fname, headers)
        elif scene.link_utils_file_ext == '1':
            fname = fname + '_linked_list.csv'
            dict_to_csv(linked_dict,path + fname, headers)
        else:
            fname = fname + '_linked_list.xlsx'
            dict_to_xlsx(linked_dict,path + fname, headers)

        return {'FINISHED'}

class PanelLinkedUtils(bpy.types.Panel):
    bl_label = "Linked Library Utilities"
    bl_space_type = "VIEW_3D"
    bl_region_type = "TOOLS"
    bl_category = "Relations"

    def draw(self, context):
        layout = self.layout
        scene = context.scene

        layout.prop(scene,'link_utils_path')
        layout.label('File extension:')
        layout.prop(scene,'link_utils_file_ext',expand=True)
        layout.operator('object.save_linked', icon='LINK_BLEND', text='Save Linked List')
        layout.prop(scene,'show_linked_list')

        box = layout.box()

        if scene.show_linked_list:

            row = box.row()
            row.label(text='Obj Name')
            row.label(text='[Path, num]')

            linked_dict = get_linked_dict()

            keys = list(linked_dict.keys())
            keys.sort()
            for key in keys:
                row = box.row()
                row.label(text=key + ':')
                row.label(text=str(linked_dict[key]))


def register():
    bpy.utils.register_class(PanelLinkedUtils)
    bpy.utils.register_class(SaveLinkedList)
    bpy.types.Scene.link_utils_path = bpy.props.StringProperty(
        name = "File Path",
        default = '//',
        description = "Define the path where the linked list file will be saved",
        subtype = 'DIR_PATH'
    )

    bpy.types.Scene.show_linked_list = bpy.props.BoolProperty(
        name = "Show Linked List",
        description = "Shows a list of linked objects"
    )

    bpy.types.Scene.link_utils_file_ext = bpy.props.EnumProperty(name='File ext',items=(('0','.txt',''),('1','.csv',''),('2','.xlsx','')))

def unregister():
    bpy.utils.unregister_class(PanelLinkedUtils)
    bpy.utils.unregister_class(SaveLinkedList)
    del bpy.types.Scene.link_utils_path
    del bpy.types.Scene.show_linked_list
    del bpy.types.Scene.link_utils_file_ext

if __name__ == "__main__":
    register()