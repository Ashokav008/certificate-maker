path = "E:\\codevert_certificate\\novice2.xlsx"

import cv2
import pandas as pd

data = pd.read_excel(path)

names=list(data.Name)

print(names)

for index, name in enumerate(names):
    template=cv2.imread('certificate.jpg')
    cv2.putText(template,name,(649,921),cv2.FONT_HERSHEY_COMPLEX,1.3,(0,0,0),3,cv2.LINE_4)
    cv2.imwrite(f'E:\\codevert_certificate\\NOVICE_2\{name}.jpg',template)
    print('processing certificate {}/{}'.format(index+1,len(names)))
    