import { Component, type ErrorInfo, type ReactNode } from 'react'

interface ErrorBoundaryProps {
  children: ReactNode
}

interface ErrorBoundaryState {
  error: Error | null
}

/**
 * 화면 전체가 새하얗게(사실은 까맣게) 죽어버리는 대신 최소한의 복구 경로를
 * 보여준다. 실제로 렌더링 중 에러가 한 번 발생했는데 아무 안내 없이 빈
 * 화면만 남는 걸 본 뒤 추가했다 — telemetry(M6)가 붙으면 여기서 리포트한다.
 */
export class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  state: ErrorBoundaryState = { error: null }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error }
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error('Unhandled render error:', error, info.componentStack)
  }

  render() {
    if (this.state.error) {
      return (
        <div className="app-empty" style={{ height: '100%' }}>
          <div className="dropzone">
            <div className="dropzone-title">문제가 생겼어요</div>
            <div className="dropzone-sub">
              화면을 그리다가 예상치 못한 오류가 났어요. 새로고침하면 대부분 해결돼요 — 파일은 서버에 없으니
              다시 열어서 이어가시면 돼요.
            </div>
            <button className="btn btn--primary" type="button" onClick={() => window.location.reload()}>
              새로고침
            </button>
          </div>
        </div>
      )
    }
    return this.props.children
  }
}
