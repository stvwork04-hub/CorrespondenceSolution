# Quick Start Guide - System User Lookup

## Option 1: Use the Complete React Component (Easiest)

```typescript
import { SystemUserLookup } from "./helpers";

// In your render method
<SystemUserLookup
  context={this.context}
  allowMultipleSelection={true}
  onSelectionChanged={(users) => {
    console.log(
      "Selected:",
      users.map((u) => u.fullname)
    );
  }}
/>;
```

## Option 2: Use the Hook for Custom UI

```typescript
import { useSystemUserLookup } from "./helpers";

const MyCustomLookup = ({ context }) => {
  const { users, loading, pagination, actions } = useSystemUserLookup({
    context,
  });

  return (
    <div>
      <input onChange={(e) => actions.search(e.target.value)} />
      {loading
        ? "Loading..."
        : users.map((user) => (
            <div key={user.systemuserid}>{user.fullname}</div>
          ))}
    </div>
  );
};
```

## Option 3: Use Just the Service (Most Control)

```typescript
import { SystemUserService } from "./helpers";

const service = new SystemUserService(context);

// Get page 1 with search
const result = await service.getActiveUsers(1, "john");
console.log("Users:", result.users);
console.log("Pagination:", result.pagination);
```

## Integration in Your PCF Control

```typescript
// In your index.ts
import * as React from "react";
import * as ReactDOM from "react-dom";
import { SystemUserLookup } from "./helpers";

public init(context, notifyOutputChanged, state, container) {
    this.context = context;
    this.container = container;
    this.renderComponent();
}

private renderComponent() {
    const element = React.createElement(SystemUserLookup, {
        context: this.context,
        onSelectionChanged: (users) => {
            // Handle selection - store users, trigger output change, etc.
        }
    });
    ReactDOM.render(element, this.container);
}
```

That's it! The components handle all the complexity of pagination, search, and Dataverse integration for you.
