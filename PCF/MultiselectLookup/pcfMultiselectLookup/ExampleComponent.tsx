import * as React from 'react';
import { Stack, Text, Separator } from '@fluentui/react';
import { SystemUserLookup, useSystemUserLookup, SystemUser } from './helpers';

interface ExampleComponentProps {
    context: ComponentFramework.Context<any>;
}

export const ExampleSystemUserComponent: React.FC<ExampleComponentProps> = ({ context }) => {
    const [selectedUsersFromComponent, setSelectedUsersFromComponent] = React.useState<SystemUser[]>([]);
    const [selectedUsersFromHook, setSelectedUsersFromHook] = React.useState<SystemUser[]>([]);

    // Example using the hook
    const {
        users,
        loading,
        error,
        pagination,
        actions
    } = useSystemUserLookup({
        context,
        pageSize: 10, // Smaller page size for demo
        autoLoad: true
    });

    const handleSelectionFromComponent = React.useCallback((users: SystemUser[]) => {
        setSelectedUsersFromComponent(users);
    }, []);

    const handleSearchFromHook = React.useCallback(async (searchTerm: string) => {
        await actions.search(searchTerm);
    }, [actions]);

    const handleUserClick = React.useCallback((user: SystemUser) => {
        const isSelected = selectedUsersFromHook.some(u => u.systemuserid === user.systemuserid);
        if (isSelected) {
            setSelectedUsersFromHook(prev => prev.filter(u => u.systemuserid !== user.systemuserid));
        } else {
            setSelectedUsersFromHook(prev => [...prev, user]);
        }
    }, [selectedUsersFromHook]);

    return (
        <Stack tokens={{ childrenGap: 20 }}>
            <Text variant="xLarge">System User Lookup Examples</Text>
            
            <Separator />
            
            {/* Example 1: Using the React Component */}
            <Stack tokens={{ childrenGap: 10 }}>
                <Text variant="large">Example 1: React Component</Text>
                <Text>Selected users: {selectedUsersFromComponent.map(u => `${u.firstname} ${u.lastname}`.trim()).join(', ')}</Text>
                <SystemUserLookup
                    context={context}
                    allowMultipleSelection={true}
                    onSelectionChanged={handleSelectionFromComponent}
                />
            </Stack>

            <Separator />

            {/* Example 2: Using the Hook */}
            <Stack tokens={{ childrenGap: 10 }}>
                <Text variant="large">Example 2: Custom Implementation with Hook</Text>
                <Text>Selected users: {selectedUsersFromHook.map(u => `${u.firstname} ${u.lastname}`.trim()).join(', ')}</Text>
                
                {error && <Text style={{ color: 'red' }}>Error: {error}</Text>}
                
                <div>
                    <input
                        type="text"
                        placeholder="Search users..."
                        onChange={(e) => handleSearchFromHook(e.target.value)}
                        style={{ marginBottom: 10, padding: 8, width: 300 }}
                    />
                </div>

                {loading ? (
                    <Text>Loading...</Text>
                ) : (
                    <div>
                        <Text>
                            Page {pagination.currentPage} of {pagination.totalPages} 
                            ({pagination.totalRecords} total records)
                        </Text>
                        
                        <div style={{ marginTop: 10 }}>
                            <button 
                                onClick={actions.previousPage} 
                                disabled={!pagination.hasPreviousPage}
                                style={{ marginRight: 10 }}
                            >
                                Previous
                            </button>
                            <button 
                                onClick={actions.nextPage} 
                                disabled={!pagination.hasNextPage}
                            >
                                Next
                            </button>
                        </div>

                        <div style={{ marginTop: 20 }}>
                            {users.map(user => (
                                <div 
                                    key={user.systemuserid}
                                    onClick={() => handleUserClick(user)}
                                    style={{
                                        padding: 10,
                                        border: '1px solid #ccc',
                                        margin: '5px 0',
                                        cursor: 'pointer',
                                        backgroundColor: selectedUsersFromHook.some(u => u.systemuserid === user.systemuserid) 
                                            ? '#e3f2fd' : 'white'
                                    }}
                                >
                                    <strong>{`${user.firstname} ${user.lastname}`.trim()}</strong><br />
                                    <small>{user.internalemailaddress}</small>
                                </div>
                            ))}
                        </div>
                    </div>
                )}
            </Stack>
        </Stack>
    );
};