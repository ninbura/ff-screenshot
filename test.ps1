ffmpeg -y -f dshow -rtbufsize 2147.48M -video_size 3440x1440 `
-i video="AVerMedia HD Capture GC573 1" -map 0 `
-pix_fmt yuvj444p -update 1 -frames:v 1 -q:v 2 `
"C:\Users\gabri\Pictures\test.jpeg"