# React（配合 TypeScript）专属最佳实践

以下仅列 LLM 写 React 时最常踩的坑，通用原则（见根目录 [CLAUDE.md](../CLAUDE.md)）不再复述。

## Hooks 与状态
- `useEffect` 是最后手段：能用派生状态（直接在渲染里算）、事件处理器、或数据获取库替代就替代。不要用 effect 来"同步 state"。
- State 保持最小集；凡是能从 props 或其他 state 推导出来的，就不要再建一个 state。
- 列表 `key` 用业务稳定 ID，**不要**用数组 index（除非列表真的静态）。
- `useMemo` / `useCallback` 只在两种情况下用：1) 子组件通过 `React.memo` 或依赖数组依赖这个引用；2) 有明确 profiler 证据。没证据就不用。

## 数据获取
- 客户端数据拉取首选 TanStack Query / SWR / 框架 loader，不要手写 `useEffect + fetch + setState`——会错过 race condition、重试、缓存、取消。
- 处理 loading / error / empty 三态；永远别假设数据已就绪。

## TypeScript
- 禁止 `any`；必须逃逸时用 `unknown` 再缩小。
- 有限选项用字符串字面量联合（`'idle' | 'loading' | 'error'`），不要 `string`。
- 变体状态用 discriminated union（`{ status: 'ok', data } | { status: 'error', err }`），不要多个可选字段并存。
- Props 里的回调显式写返回类型 `() => void`，避免意外吃掉 Promise。

## 组件设计
- 受控 vs 非受控，组件内二选一，不要混用。
- 跨层共享状态：先考虑提升 + 组合 / children 模式，再考虑 Context；Context 只放"真正全局"的东西（主题、当前用户）。
- 避免在渲染中 mutate：state/props 都是不可变；数组/对象更新返回新引用。
- 用 `useRef` 存"不触发渲染的可变值"，别用 `useState`。

## Next.js / RSC（若适用）
- 清楚每个文件是 Server Component 还是 Client Component（顶行 `"use client"`）；不要把服务端才能做的事放到客户端。
- 数据获取优先放在 Server Component / Route Handler / Server Action。
