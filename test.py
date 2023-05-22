import os
import yaml

data = {"label1": "Value 1", "label2": 12345, "label3": ["a", "b", "c"], "label4": "Value 4",}

yaml_string = yaml.dump(data)
print(yaml_string)

with open("ex.yml", mode='w') as f:
    f.write(yaml_string)

with open("ex.yml", mode='r') as rf:
    data2 = yaml.safe_load(rf)
print(data2)
