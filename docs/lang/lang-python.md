# Python 专属最佳实践

以下仅列 LLM 写 Python 时最常踩的坑，通用原则（见根目录 [CLAUDE.md](../CLAUDE.md)）不再复述。

## 类型与数据结构
- 所有公共函数/方法必须有类型注解；用 `mypy` 或 `pyright` 把关。
- 结构化数据用 `@dataclass(slots=True)` / `pydantic.BaseModel` / `TypedDict`，不要到处传 `dict`。
- `pathlib.Path` 替代 `os.path`；`datetime` 始终 timezone-aware（`datetime.now(timezone.utc)`）。

## 常见陷阱
- **可变默认参数**：禁止 `def f(x=[])` / `def f(x={})`；用 `None` 哨兵。
- 不要广捕 `except Exception`（更不要 `except:`）；捕具体异常，或至少 `logger.exception` 后重抛。
- 迭代大数据用生成器/迭代器，避免把大列表一次性 materialize。
- 修改传入参数属于副作用，默认返回新对象；必须原地改时函数名体现（如 `sort` vs `sorted`）。

## I/O 与资源
- 文件、锁、连接一律用 `with` 上下文管理器。
- 库代码用 `logging`，不用 `print`；日志里不要拼 f-string 作为 msg，用 `logger.info("x=%s", x)` 以便 lazy formatting。

## 工程规范
- 字符串格式化统一 f-string；旧式 `%` / `.format` 只在日志 lazy 场景保留。
- 依赖锁定（`uv.lock` / `poetry.lock` / `requirements.txt` 带 hash）；虚拟环境隔离。
- 格式化和 lint 用 `ruff format` + `ruff check`；避免 `black` + `isort` + `flake8` 多工具拼接。
- 避免 `from module import *`。
