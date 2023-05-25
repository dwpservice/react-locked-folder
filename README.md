Full code for React locked folder project on App Script ID 1kgbJuxdXsP-uMhRfkXcWnGRKEn0e8uJ9QVkX-ZEGSlBu0j5AB3ObruTP.

Log in with Clasp and redownload Node_modules.

Run npm start to dev locally. Run npm run build for build version to use with App Script webapp.

Use npm run gstart to watch for change in folder apps-script and push back to Apps Script automatically or use npm run gpush to manually push once at a time. Please make sure it is up to date with App Script version (using npm run gpull).

```
/* clone github repo to local or fetch the data */

git init
git remote add origin https://github.com/goatshub/react-locked-folder.git
git fetch --all
git pull origin main
```

```
/* run project */

npm i

npm run glogin //log in to clasp with driveadmin email (install/init clasp if needed)

npm run gpull //to update from google script project first before pushing

npm run build //build react frontend into html and overwrite index.html in apps-script folder

npm run gpush //push back to app script
//or
npm run gstart //automatically watch the change in apps-script and push

```
