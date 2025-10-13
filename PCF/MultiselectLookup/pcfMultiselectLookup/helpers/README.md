# System User Lookup Helper Components

This directory contains helper components for creating a system user lookup with pagination and search functionality for Power Apps Component Framework (PCF) controls.

## Components Overview

### 1. SystemUserService (`SystemUserService.ts`)

A service class that handles all interactions with the Dataverse SystemUser table.

**Features:**

- Retrieves active system users with pagination (25 records per page by default)
- Supports fuzzy search on firstname, lastname, and fullname
- Configurable page size
- Methods for getting single user or multiple users by ID

**Usage:**

```typescript
import { SystemUserService } from "./helpers";

const userService = new SystemUserService(context);

// Get users with pagination and search
const result = await userService.getActiveUsers(1, "john");

// Get user by ID
const user = await userService.getUserById("user-guid");

// Get multiple users by IDs
const users = await userService.getUsersByIds(["guid1", "guid2"]);
```

### 2. SystemUserLookup Component (`SystemUserLookup.tsx`)

A complete React component with search, pagination, and selection capabilities.

**Features:**

- Search box with debounced search (500ms delay)
- Pagination controls (First, Previous, Next, Last)
- Single or multiple selection modes
- Loading states and error handling
- Fluent UI components for consistent styling

**Usage:**

```typescript
import { SystemUserLookup } from "./helpers";

<SystemUserLookup
  context={context}
  allowMultipleSelection={true}
  onSelectionChanged={(selectedUsers) => {
    console.log("Selected users:", selectedUsers);
  }}
  initialSelectedUsers={[]}
/>;
```

### 3. useSystemUserLookup Hook (`useSystemUserLookup.ts`)

A React hook for building custom user lookup interfaces.

**Features:**

- State management for users, loading, pagination
- Action methods for loading, searching, navigation
- Customizable page size and auto-load behavior

**Usage:**

```typescript
import { useSystemUserLookup } from "./helpers";

const MyComponent = ({ context }) => {
  const { users, loading, error, pagination, actions } = useSystemUserLookup({
    context,
    pageSize: 10,
    autoLoad: true,
  });

  return (
    <div>
      <input
        onChange={(e) => actions.search(e.target.value)}
        placeholder="Search users..."
      />
      {users.map((user) => (
        <div key={user.systemuserid}>{user.fullname}</div>
      ))}
      <button
        onClick={actions.previousPage}
        disabled={!pagination.hasPreviousPage}
      >
        Previous
      </button>
      <button onClick={actions.nextPage} disabled={!pagination.hasNextPage}>
        Next
      </button>
    </div>
  );
};
```

## Data Structure

### SystemUser Interface

```typescript
interface SystemUser {
  systemuserid: string;
  firstname: string;
  lastname: string;
  fullname: string;
  internalemailaddress: string;
  businessunitid: string;
  title: string;
}
```

### PaginationInfo Interface

```typescript
interface PaginationInfo {
  currentPage: number;
  pageSize: number;
  totalRecords: number;
  totalPages: number;
  hasNextPage: boolean;
  hasPreviousPage: boolean;
}
```

## Installation and Setup

1. **Import the components** in your PCF control:

```typescript
import {
  SystemUserService,
  SystemUserLookup,
  useSystemUserLookup,
} from "./helpers";
```

2. **For React components**, make sure your PCF project has React dependencies:

```json
{
  "dependencies": {
    "@fluentui/react": "^8.121.1",
    "react": "^16.14.0",
    "react-dom": "^16.14.0"
  }
}
```

3. **Initialize in your PCF control**:

```typescript
// In your init method
private renderComponent(): void {
    const reactElement = React.createElement(SystemUserLookup, {
        context: this.context,
        onSelectionChanged: (users) => {
            // Handle selection
        }
    });
    ReactDOM.render(reactElement, this.container);
}
```

## Search Functionality

The search feature supports fuzzy matching on:

- First Name (firstname)
- Last Name (lastname)
- Full Name (fullname)

Search is case-insensitive and uses the OData `contains` function for partial matching.

## Pagination

- Default page size: 25 records
- Configurable page size via service or hook options
- Navigation: First, Previous, Next, Last
- Page info display: "Page X of Y (Z total records)"

## Error Handling

All components include comprehensive error handling:

- Network errors from Dataverse calls
- Invalid pagination parameters
- Search query errors
- User-friendly error messages displayed in UI

## Performance Considerations

- **Debounced Search**: 500ms delay to prevent excessive API calls
- **Efficient Queries**: Only retrieves necessary fields
- **Count Optimization**: Separate count query for pagination
- **Memory Management**: Proper cleanup in React components

## Integration Example

See `ExampleComponent.tsx` for a complete working example that demonstrates:

- Both component and hook usage patterns
- Selection handling
- Search implementation
- Pagination controls
- Error states

## Customization

### Custom Fields

To add additional fields, modify the SystemUser interface and update the OData select statements in SystemUserService.

### Custom Styling

The components use Fluent UI by default. You can:

- Override Fluent UI themes
- Replace with custom components
- Use the hook to build completely custom UI

### Custom Filtering

Extend SystemUserService to add additional filters:

```typescript
// Add business unit filter
private buildBusinessUnitFilter(businessUnitId: string): string {
    return `_businessunitid_value eq '${businessUnitId}'`;
}
```
