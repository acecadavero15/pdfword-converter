1. Put these 4 files (Dockerfile, package.json, server.js, README.txt) in a NEW GitHub repository.
2. On render.com: New > Web Service > connect that repository > Runtime: Docker > Instance type: Free > Create.
3. When it says "Live", copy the address (https://something.onrender.com).
4. Open script.js of the website and set:  const CONVERTER_URL = 'https://something.onrender.com';
5. Re-upload script.js to the website repository.
Optional: add an environment variable ALLOWED_ORIGIN = https://YOURNAME.github.io so only your site can use the server.
