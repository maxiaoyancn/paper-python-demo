# Golang 专属最佳实践

以下仅列 LLM 写 Go 时最常踩的坑，通用原则（见根目录 [CLAUDE.md](../CLAUDE.md)）不再复述。

## 错误处理
- 每一个 `err` 都要检查，不允许 `_ = fn()`。
- 包装用 `fmt.Errorf("op xxx: %w", err)` 保留链；判等用 `errors.Is` / `errors.As`，不要 `==`。
- 不要在库代码里 `panic`；只有真正无法恢复（配置缺失、启动失败）才用。

## 并发
- 启协程前先想清楚"谁负责它的生命周期"；用 `errgroup.Group` 或 `sync.WaitGroup` 收束，不留孤儿。
- Channel：发送方负责 `close`，接收方不关；ranges 用 `for v := range ch`。
- `context.Context` 放第一个参数，向下传递，不要塞进 struct；收到 `ctx.Done()` 必须尽快返回。

## API 设计
- 接受接口，返回结构体。接口定义在"使用方"包里，不在"实现方"包里。
- 让零值可用（`var x Foo` 就能工作），优先级高于加构造函数。
- 不要预先加泛型；等到第二个具体用例出现再抽。

## 工程规范
- 包名小写、单数、有意义；避免 `util`、`common`、`helper`。
- 避免 `init()`；启动逻辑放 `main` 或显式 `Setup()`。
- 测试首选表驱动 + `t.Run` 子测试；辅助函数首行 `t.Helper()`。
- 提交前跑 `go vet` 和 `golangci-lint run`。
