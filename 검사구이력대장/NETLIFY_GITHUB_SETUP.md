# Netlify + GitHub 자동 배포 설정

이 프로젝트는 `export` 폴더를 정적 사이트로 배포합니다.

## Netlify 설정

1. GitHub에 `검사구이력대장` 폴더를 포함해서 push합니다.
2. Netlify에서 `Add new site` -> `Import an existing project` -> GitHub 저장소를 선택합니다.
3. 저장소 루트가 `검사구이력대장`이면 그대로 진행합니다.
4. 저장소 루트가 `C:\Python` 전체라면 Netlify의 `Base directory`를 `검사구이력대장`으로 지정합니다.
5. `Build command`는 비워둡니다.
6. `Publish directory`는 `export`로 지정합니다.

## 앱 설정

검사구 프로그램의 설정에서 `Netlify URL`에 Netlify 사이트 주소를 입력합니다.

예:

```text
https://sejiqc.netlify.app
```

이후 `전체 HTML/QR 갱신`을 실행하면 QR 코드가 Netlify 주소 기준으로 생성됩니다.

## 배포 흐름

1. 검사구 프로그램에서 HTML/QR을 갱신합니다.
2. 변경된 `export/index.html`, `export/cards/*.html`, `export/qrcode/*.png`를 GitHub에 push합니다.
3. Netlify가 GitHub 변경을 감지해 자동 배포합니다.

NAS 포트포워딩, DDNS, WebDAV 설정은 Netlify 방식에서는 비워둬도 됩니다.
