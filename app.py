from win32com.client.dynamic import Dispatch

zulu_map_doc = Dispatch("ZuluLib.MapDoc")

if zulu_map_doc.Open("D:\\Study\\Zulu_plugin\\test_data\\map.zmp"):
    layers = zulu_map_doc.Layers
    for i in range(1, layers.Count+1):
        if layers.Item(i).UserName == "test_1":
            layer = layers.Item(i)
types = layer.ObjectTypes
for i in range(1, types.Count+1):
    l_type = types.Item(i)
    print(l_type.Name)
    for j in range(1,l_type.Modes.Count+1):
        mode = l_type.Modes.Item(j)
        print("    "+mode.Name)

