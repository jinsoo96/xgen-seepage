"""xgen-seepage 자체 XGEN 서버 커넥터.

xgen-connector(PlateerLab, Electron)에 의존하지 않는다 — 설치돼 있지 않아도
동작해야 한다는 요구사항 때문에, 로그인/토큰관리/에이전트-도구 브릿지를
전부 이 프로젝트 안에서 독립적으로 구현한다. xgen-connector의 `src/core/`
(auth.ts/hash.ts/client.ts)와 `src/main/mcp-bridge.ts`는 **와이어 프로토콜을
파악하기 위한 참고 자료로만** 썼다 — 코드를 가져오지 않고 같은 프로토콜을
Python으로 새로 짰다. 자세한 근거는 저장소 `NOTICE`.
"""
