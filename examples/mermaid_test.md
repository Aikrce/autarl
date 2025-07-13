# Mermaid图表测试文档

本文档用于测试Mermaid图表到Word的转换功能。

## 1. 流程图 (Flowchart)

下面是一个简单的流程图示例：

```mermaid
flowchart TD
    A[开始] --> B{是否登录?}
    B -->|是| C[进入主页]
    B -->|否| D[显示登录页面]
    D --> E[输入用户名密码]
    E --> F{验证通过?}
    F -->|是| C
    F -->|否| D
    C --> G[结束]
```

## 2. 序列图 (Sequence Diagram)

用户认证流程序列图：

```mermaid
sequenceDiagram
    participant U as 用户
    participant C as 客户端
    participant S as 服务器
    participant DB as 数据库
    
    U->>C: 输入用户名密码
    C->>S: 发送登录请求
    S->>DB: 查询用户信息
    DB-->>S: 返回用户数据
    S-->>C: 返回认证结果
    C-->>U: 显示登录状态
```

## 3. 类图 (Class Diagram)

系统核心类结构：

```mermaid
classDiagram
    class User {
        +String id
        +String name
        +String email
        +login()
        +logout()
    }
    
    class Order {
        +String orderId
        +Date createTime
        +Float amount
        +create()
        +cancel()
    }
    
    class Product {
        +String productId
        +String name
        +Float price
        +getDetails()
    }
    
    User "1" --> "*" Order : places
    Order "*" --> "*" Product : contains
```

## 4. 状态图 (State Diagram)

订单状态转换图：

```mermaid
stateDiagram-v2
    [*] --> 待支付
    待支付 --> 已支付: 支付成功
    待支付 --> 已取消: 取消订单
    已支付 --> 配送中: 开始配送
    配送中 --> 已完成: 确认收货
    已完成 --> [*]
    已取消 --> [*]
```

## 5. 实体关系图 (ER Diagram)

数据库设计：

```mermaid
erDiagram
    CUSTOMER {
        int id PK
        string name
        string email
        string phone
    }
    
    ORDER {
        int id PK
        int customer_id FK
        date order_date
        float total_amount
    }
    
    PRODUCT {
        int id PK
        string name
        float price
        int stock
    }
    
    ORDER_ITEM {
        int order_id FK
        int product_id FK
        int quantity
        float price
    }
    
    CUSTOMER ||--o{ ORDER : places
    ORDER ||--|{ ORDER_ITEM : contains
    PRODUCT ||--o{ ORDER_ITEM : "ordered in"
```

## 6. 甘特图 (Gantt Chart)

项目计划时间线：

```mermaid
gantt
    title 项目开发计划
    dateFormat  YYYY-MM-DD
    section 需求分析
    需求调研           :a1, 2024-01-01, 7d
    需求文档编写       :after a1, 5d
    
    section 设计阶段
    系统设计           :2024-01-13, 10d
    数据库设计         :2024-01-15, 7d
    
    section 开发阶段
    后端开发           :2024-01-23, 20d
    前端开发           :2024-01-25, 18d
    
    section 测试阶段
    单元测试           :2024-02-10, 7d
    集成测试           :2024-02-15, 5d
    上线部署           :2024-02-20, 2d
```

## 7. 饼图 (Pie Chart)

市场份额分布：

```mermaid
pie title 2024年市场份额
    "产品A" : 35
    "产品B" : 25
    "产品C" : 20
    "产品D" : 10
    "其他" : 10
```

## 8. 用户旅程图 (Journey)

用户购物体验旅程：

```mermaid
journey
    title 用户网购体验旅程
    section 浏览商品
      浏览首页: 5: 用户
      搜索商品: 4: 用户
      查看详情: 5: 用户
    section 下单购买
      加入购物车: 5: 用户
      填写地址: 3: 用户
      选择支付: 4: 用户
    section 售后服务
      收到商品: 5: 用户
      申请退货: 2: 用户
      客服处理: 4: 用户
```

## 9. Git图 (Git Graph)

版本控制流程：

```mermaid
gitGraph
    commit
    commit
    branch develop
    checkout develop
    commit
    commit
    checkout main
    merge develop
    commit
    branch feature
    checkout feature
    commit
    commit
    checkout develop
    merge feature
    checkout main
    merge develop
```

## 总结

以上展示了Mermaid支持的主要图表类型：

1. **流程图** - 用于展示流程和决策逻辑
2. **序列图** - 用于展示时序交互
3. **类图** - 用于展示类结构和关系
4. **状态图** - 用于展示状态转换
5. **ER图** - 用于数据库设计
6. **甘特图** - 用于项目计划
7. **饼图** - 用于数据占比展示
8. **旅程图** - 用于用户体验分析
9. **Git图** - 用于版本控制流程

这些图表在转换到Word文档后，应该以图片形式正确显示，并保持良好的可读性。