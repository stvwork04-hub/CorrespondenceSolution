import * as React from 'react';
import {
    DetailsList,
    IColumn,
    SelectionMode,
    SearchBox,
    Stack,
    Text,
    IconButton,
    CommandBar,
    ICommandBarItemProps,
    MessageBar,
    MessageBarType,
    Spinner,
    SpinnerSize
} from '@fluentui/react';
import { SystemUserService, SystemUser, SystemUserSearchResult } from './SystemUserService';

interface SystemUserLookupProps {
    context: ComponentFramework.Context<any>;
    onSelectionChanged?: (selectedUsers: SystemUser[]) => void;
    allowMultipleSelection?: boolean;
    initialSelectedUsers?: SystemUser[];
}

interface SystemUserLookupState {
    users: SystemUser[];
    loading: boolean;
    error: string | null;
    searchTerm: string;
    currentPage: number;
    totalPages: number;
    totalRecords: number;
    hasNextPage: boolean;
    hasPreviousPage: boolean;
    selectedUsers: SystemUser[];
}

export class SystemUserLookup extends React.Component<SystemUserLookupProps, SystemUserLookupState> {
    private userService: SystemUserService;
    private searchTimeout: NodeJS.Timeout | null = null;

    constructor(props: SystemUserLookupProps) {
        super(props);
        
        this.userService = new SystemUserService(props.context);
        
        this.state = {
            users: [],
            loading: false,
            error: null,
            searchTerm: '',
            currentPage: 1,
            totalPages: 0,
            totalRecords: 0,
            hasNextPage: false,
            hasPreviousPage: false,
            selectedUsers: props.initialSelectedUsers || []
        };
    }

    public async componentDidMount(): Promise<void> {
        await this.loadUsers();
    }

    private async loadUsers(page: number = 1, searchTerm?: string): Promise<void> {
        this.setState({ loading: true, error: null });

        try {
            const result: SystemUserSearchResult = await this.userService.getActiveUsers(page, searchTerm);
            
            this.setState({
                users: result.users,
                currentPage: result.pagination.currentPage,
                totalPages: result.pagination.totalPages,
                totalRecords: result.pagination.totalRecords,
                hasNextPage: result.pagination.hasNextPage,
                hasPreviousPage: result.pagination.hasPreviousPage,
                loading: false
            });

        } catch (error) {
            this.setState({
                error: `Failed to load users: ${error}`,
                loading: false
            });
        }
    }

    private onSearchChange = (_?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
        const searchTerm = newValue || '';
        this.setState({ searchTerm });

        // Clear existing timeout
        if (this.searchTimeout) {
            clearTimeout(this.searchTimeout);
        }

        // Set new timeout for debounced search
        this.searchTimeout = setTimeout(() => {
            this.loadUsers(1, searchTerm);
        }, 500);
    };

    private onNextPage = (): void => {
        if (this.state.hasNextPage) {
            this.loadUsers(this.state.currentPage + 1, this.state.searchTerm);
        }
    };

    private onPreviousPage = (): void => {
        if (this.state.hasPreviousPage) {
            this.loadUsers(this.state.currentPage - 1, this.state.searchTerm);
        }
    };

    private onGoToFirstPage = (): void => {
        this.loadUsers(1, this.state.searchTerm);
    };

    private onGoToLastPage = (): void => {
        this.loadUsers(this.state.totalPages, this.state.searchTerm);
    };

    private onSelectionChanged = (): void => {
        // This will be handled by the DetailsList selection
        if (this.props.onSelectionChanged) {
            this.props.onSelectionChanged(this.state.selectedUsers);
        }
    };

    private onActiveItemChanged = (item?: SystemUser, index?: number): void => {
        if (!item) return;
        
        if (this.props.allowMultipleSelection) {
            // Handle multi-select
            const isSelected = this.state.selectedUsers.some(u => u.systemuserid === item.systemuserid);
            let newSelection: SystemUser[];
            
            if (isSelected) {
                newSelection = this.state.selectedUsers.filter(u => u.systemuserid !== item.systemuserid);
            } else {
                newSelection = [...this.state.selectedUsers, item];
            }
            
            this.setState({ selectedUsers: newSelection });
            
            if (this.props.onSelectionChanged) {
                this.props.onSelectionChanged(newSelection);
            }
        } else {
            // Handle single select
            const newSelection = [item];
            this.setState({ selectedUsers: newSelection });
            
            if (this.props.onSelectionChanged) {
                this.props.onSelectionChanged(newSelection);
            }
        }
    };

    private getColumns(): IColumn[] {
        return [
            {
                key: 'fullname',
                name: 'Full Name',
                fieldName: 'fullname',
                minWidth: 150,
                maxWidth: 200,
                isResizable: true,
                onRender: (item: SystemUser) => (
                    <Text>{`${item.firstname} ${item.lastname}`.trim()}</Text>
                )
            },
            {
                key: 'firstname',
                name: 'First Name',
                fieldName: 'firstname',
                minWidth: 100,
                maxWidth: 150,
                isResizable: true
            },
            {
                key: 'lastname',
                name: 'Last Name',
                fieldName: 'lastname',
                minWidth: 100,
                maxWidth: 150,
                isResizable: true
            },
            {
                key: 'internalemailaddress',
                name: 'Email',
                fieldName: 'internalemailaddress',
                minWidth: 200,
                maxWidth: 250,
                isResizable: true
            }
        ];
    }

    private getCommandBarItems(): ICommandBarItemProps[] {
        return [
            {
                key: 'firstPage',
                text: 'First',
                iconProps: { iconName: 'DoubleChevronLeft' },
                disabled: !this.state.hasPreviousPage || this.state.loading,
                onClick: this.onGoToFirstPage
            },
            {
                key: 'previousPage',
                text: 'Previous',
                iconProps: { iconName: 'ChevronLeft' },
                disabled: !this.state.hasPreviousPage || this.state.loading,
                onClick: this.onPreviousPage
            },
            {
                key: 'pageInfo',
                text: `Page ${this.state.currentPage} of ${this.state.totalPages} (${this.state.totalRecords} total)`,
                disabled: true
            },
            {
                key: 'nextPage',
                text: 'Next',
                iconProps: { iconName: 'ChevronRight' },
                disabled: !this.state.hasNextPage || this.state.loading,
                onClick: this.onNextPage
            },
            {
                key: 'lastPage',
                text: 'Last',
                iconProps: { iconName: 'DoubleChevronRight' },
                disabled: !this.state.hasNextPage || this.state.loading,
                onClick: this.onGoToLastPage
            }
        ];
    }

    public render(): React.ReactElement {
        const { loading, error, users, selectedUsers } = this.state;

        return (
            <Stack tokens={{ childrenGap: 10 }}>
                {/* Search Box */}
                <SearchBox
                    placeholder="Search by name..."
                    value={this.state.searchTerm}
                    onChange={this.onSearchChange}
                    disabled={loading}
                />

                {/* Error Message */}
                {error && (
                    <MessageBar messageBarType={MessageBarType.error}>
                        {error}
                    </MessageBar>
                )}

                {/* Selected Users Info */}
                {selectedUsers.length > 0 && (
                    <MessageBar messageBarType={MessageBarType.info}>
                        {selectedUsers.length} user(s) selected: {selectedUsers.map(u => `${u.firstname} ${u.lastname}`.trim()).join(', ')}
                    </MessageBar>
                )}

                {/* Loading Spinner */}
                {loading && (
                    <Stack horizontalAlign="center">
                        <Spinner size={SpinnerSize.large} label="Loading users..." />
                    </Stack>
                )}

                {/* Users List */}
                {!loading && users.length > 0 && (
                    <>
                        <DetailsList
                            items={users}
                            columns={this.getColumns()}
                            selectionMode={this.props.allowMultipleSelection ? SelectionMode.multiple : SelectionMode.single}
                            onActiveItemChanged={this.onActiveItemChanged}
                            compact={true}
                        />

                        {/* Pagination Controls */}
                        <CommandBar items={this.getCommandBarItems()} />
                    </>
                )}

                {/* No Results */}
                {!loading && users.length === 0 && !error && (
                    <MessageBar messageBarType={MessageBarType.info}>
                        No users found.
                    </MessageBar>
                )}
            </Stack>
        );
    }

    public componentWillUnmount(): void {
        if (this.searchTimeout) {
            clearTimeout(this.searchTimeout);
        }
    }
}