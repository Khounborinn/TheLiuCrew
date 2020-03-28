import pyrealsense2 as rs

import math as m

pipe = rs.pipeline()
profile = pipe.start()
try:
    for i in range(0, 100):
        frames = pipe.wait_for_frames()
        for f in frames:
            print(f.profile)
finally:
    pipe.stop()
