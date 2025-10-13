import * as React from 'react';
import {
    Stack,
    TextField,
    DefaultButton,
    Panel,
    PanelType,
    Text,
    MessageBar,
    MessageBarType,
    PrimaryButton,
    Separator
} from '@fluentui/react';
import { SystemUser, SystemUserService } from '../helpers';

// Add CSS for loading spinner animation
const spinnerStyles = `
@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}
`;

// Inject styles into document head if not already present
if (typeof document !== 'undefined' && !document.getElementById('pcf-spinner-styles')) {
    const styleSheet = document.createElement('style');
    styleSheet.id = 'pcf-spinner-styles';
    styleSheet.textContent = spinnerStyles;
    document.head.appendChild(styleSheet);
}

export interface SystemUserForwardComponentProps {
    context: ComponentFramework.Context<any>;
    initialEmails?: string;
    onEmailsChanged?: (emails: string) => void;
    disabled?: boolean;
}

export interface SystemUserForwardComponentState {
    selectedEmails: string;
    isModalOpen: boolean;
    error: string | null;
}

export class SystemUserForwardComponent extends React.Component<
    SystemUserForwardComponentProps,
    SystemUserForwardComponentState
> {
    constructor(props: SystemUserForwardComponentProps) {
        super(props);

        this.state = {
            selectedEmails: props.initialEmails || '',
            isModalOpen: false,
            error: null
        };
    }

    private onOpenModal = (): void => {
        this.setState({ isModalOpen: true, error: null });
    };

    private onCloseModal = (): void => {
        this.setState({ isModalOpen: false });
    };

    private onUsersSelected = (selectedUsers: SystemUser[]): void => {
        // Extract email addresses and join with semicolons
        const emails = selectedUsers
            .map(user => user.internalemailaddress)
            .filter(email => email && email.trim()) // Filter out empty emails
            .join(';');

        this.setState({ 
            selectedEmails: emails,
            isModalOpen: false,
            error: null 
        });

        // Notify parent component of the change
        if (this.props.onEmailsChanged) {
            this.props.onEmailsChanged(emails);
        }
    };

    private onEmailsTextChanged = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        const emails = newValue || '';
        this.setState({ selectedEmails: emails });

        if (this.props.onEmailsChanged) {
            this.props.onEmailsChanged(emails);
        }
    };

    private onClearEmails = (): void => {
        this.setState({ selectedEmails: '' });
        
        if (this.props.onEmailsChanged) {
            this.props.onEmailsChanged('');
        }
    };

    public render(): React.ReactElement {
        const { context, disabled } = this.props;
        const { selectedEmails, isModalOpen, error } = this.state;

        return (
            <Stack tokens={{ childrenGap: 15 }}>
                {/* Error Message */}
                {error && (
                    <MessageBar messageBarType={MessageBarType.error}>
                        {error}
                    </MessageBar>
                )}

                {/* Main Input Area */}
                <Stack tokens={{ childrenGap: 20 }} 
                       styles={{ 
                           root: { 
                               background: 'linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%)',
                               padding: '24px',
                               borderRadius: '12px',
                               border: '1px solid #dee2e6',
                               boxShadow: '0 4px 12px rgba(0,0,0,0.05)'
                           } 
                       }}>
                    
                    {/* Header Section with Primary Button */}
                    <Stack tokens={{ childrenGap: 16 }}>
                        <Stack horizontal horizontalAlign="start">
                            <PrimaryButton
                                text="Select users to forward"
                                iconProps={{ iconName: 'People' }}
                                onClick={this.onOpenModal}
                                disabled={disabled}
                                styles={{
                                    root: { 
                                        minWidth: '200px',
                                        height: '44px',
                                        borderRadius: '8px',
                                        backgroundColor: '#0078d4',
                                        border: 'none',
                                        fontWeight: 600,
                                        fontSize: '14px',
                                        boxShadow: '0 2px 8px rgba(0,120,212,0.3)',
                                        transition: 'all 0.2s ease'
                                    },
                                    rootHovered: {
                                        backgroundColor: '#106ebe',
                                        transform: 'translateY(-1px)',
                                        boxShadow: '0 4px 12px rgba(0,120,212,0.4)'
                                    },
                                    rootPressed: {
                                        transform: 'translateY(0px)'
                                    }
                                }}
                            />
                        </Stack>
                        
                        {/* Clear Button Row */}
                        {selectedEmails && (
                            <Stack horizontal horizontalAlign="end">
                                <DefaultButton
                                    text="Clear All"
                                    iconProps={{ iconName: 'ClearFormatting' }}
                                    onClick={this.onClearEmails}
                                    disabled={disabled}
                                    styles={{
                                        root: { 
                                            minWidth: '100px',
                                            height: '36px',
                                            borderRadius: '6px',
                                            border: '2px solid #dc3545',
                                            color: '#dc3545',
                                            backgroundColor: 'transparent',
                                            fontWeight: 500
                                        },
                                        rootHovered: {
                                            backgroundColor: '#dc3545',
                                            color: 'white',
                                            transform: 'translateY(-1px)'
                                        }
                                    }}
                                />
                            </Stack>
                        )}
                    </Stack>
                    
                    {/* Email Text Area with Enhanced Styling */}
                    <Stack tokens={{ childrenGap: 8 }}>
                        <Text variant="small" 
                              styles={{ 
                                  root: { 
                                      color: '#6c757d', 
                                      fontWeight: 500,
                                      textTransform: 'uppercase',
                                      letterSpacing: '0.5px',
                                      fontSize: '12px'
                                  } 
                              }}>
                            Selected Email Addresses
                        </Text>
                        <TextField
                            placeholder="Selected email addresses will appear here..."
                            value={selectedEmails}
                            onChange={this.onEmailsTextChanged}
                            multiline
                            rows={4}
                            disabled={disabled}
                            description="Email addresses are automatically separated by semicolons (;)"
                            styles={{
                                root: { width: '100%' },
                                fieldGroup: { 
                                    borderRadius: '8px', 
                                    border: '2px solid #e9ecef',
                                    backgroundColor: 'white',
                                    boxShadow: 'inset 0 2px 4px rgba(0,0,0,0.05)',
                                    transition: 'all 0.2s ease'
                                },

                                field: { 
                                    padding: '12px 16px',
                                    fontSize: '14px',
                                    lineHeight: '1.5',
                                    fontFamily: 'Segoe UI, sans-serif'
                                },
                                description: {
                                    color: '#6c757d',
                                    fontSize: '12px',
                                    fontStyle: 'italic',
                                    marginTop: '6px'
                                }
                            }}
                        />
                    </Stack>
                    
                    {/* Selection Counter */}
                    {selectedEmails && (
                        <Stack horizontal horizontalAlign="space-between" 
                               styles={{ 
                                   root: { 
                                       padding: '12px 16px',
                                       backgroundColor: 'rgba(0,120,212,0.1)',
                                       borderRadius: '6px',
                                       border: '1px solid rgba(0,120,212,0.2)'
                                   } 
                               }}>
                            <Text variant="small" styles={{ root: { color: '#0078d4', fontWeight: 600 } }}>
                                📧 {selectedEmails.split(';').filter(email => email.trim()).length} email(s) selected
                            </Text>
                            <Text variant="small" styles={{ root: { color: '#6c757d' } }}>
                                Ready to forward
                            </Text>
                        </Stack>
                    )}
                </Stack>

                {/* Selected Users Count */}
                {selectedEmails && (
                    <Text variant="small">
                        {selectedEmails.split(';').filter(email => email.trim()).length} recipient(s) selected
                    </Text>
                )}

                {/* Modal Dialog */}
                <Panel
                    headerText="� Select Active AD Users to Forward To"
                    isOpen={isModalOpen}
                    onDismiss={this.onCloseModal}
                    type={PanelType.medium}
                    isBlocking={true}
                    closeButtonAriaLabel="Close"
                    styles={{
                        main: { 
                            background: 'linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%)',
                        },
                        header: {
                            background: 'linear-gradient(90deg, #0078d4 0%, #106ebe 100%)',
                            color: 'white',
                            paddingTop: '20px',
                            paddingBottom: '20px'
                        },
                        headerText: {
                            color: 'white',
                            fontSize: '18px',
                            fontWeight: 700,
                            letterSpacing: '0.5px'
                        },
                        content: {
                            paddingLeft: '0px',
                            paddingRight: '0px'
                        }
                    }}
                >
                    <SystemUserSearchContent
                        context={context}
                        onUsersSelected={this.onUsersSelected}
                        onCancel={this.onCloseModal}
                        initialSelectedEmails={selectedEmails}
                    />
                </Panel>
            </Stack>
        );
    }
}

// Inline Search Content Component to avoid circular dependencies
interface SystemUserSearchContentProps {
    context: ComponentFramework.Context<any>;
    onUsersSelected: (users: SystemUser[]) => void;
    onCancel: () => void;
    initialSelectedEmails?: string;
}

interface SystemUserSearchContentState {
    searchTerm: string;
    searchResults: SystemUser[];
    selectedUsers: SystemUser[];
    loading: boolean;
    error: string | null;
    hasSearched: boolean;
    currentPage: number;
    totalPages: number;
    totalRecords: number;
    hasNextPage: boolean;
    hasPreviousPage: boolean;
}

class SystemUserSearchContent extends React.Component<
    SystemUserSearchContentProps,
    SystemUserSearchContentState
> {
    private userService: SystemUserService;

    constructor(props: SystemUserSearchContentProps) {
        super(props);

        this.userService = new SystemUserService(props.context);
        this.userService.setPageSize(25); // Show 25 records as requested

        // Pre-populate selected users based on initial emails
        const initialSelectedUsers = this.parseEmailsToUsers(props.initialSelectedEmails || '');

        this.state = {
            searchTerm: '',
            searchResults: [],
            selectedUsers: initialSelectedUsers,
            loading: false,
            error: null,
            hasSearched: false,
            currentPage: 1,
            totalPages: 1,
            totalRecords: 0,
            hasNextPage: false,
            hasPreviousPage: false
        };
    }

    public async componentDidMount(): Promise<void> {
        // Load first 25 users automatically when component mounts
        await this.loadInitialUsers();
    }

    private loadInitialUsers = async (): Promise<void> => {
        this.setState({ loading: true, error: null });

        try {
            // Clear any cached pagination data
            this.userService.clearCache();
            
            // Load first 25 users without search filter
            const result = await this.userService.getActiveUsers(1);

            this.setState({
                searchResults: result.users,
                currentPage: result.pagination.currentPage,
                totalPages: result.pagination.totalPages,
                totalRecords: result.pagination.totalRecords,
                hasNextPage: result.pagination.hasNextPage,
                hasPreviousPage: result.pagination.hasPreviousPage,
                loading: false,
                hasSearched: true
            });

        } catch (error) {
            this.setState({
                loading: false,
                error: `Failed to load users: ${error}`,
                hasSearched: true
            });
        }
    };

    private parseEmailsToUsers(emails: string): SystemUser[] {
        if (!emails) return [];

        return emails
            .split(';')
            .map(email => email.trim())
            .filter(email => email)
            .map(email => ({
                systemuserid: '',
                firstname: '',
                lastname: '',
                internalemailaddress: email
            }));
    }

    private onSearch = async (page: number = 1): Promise<void> => {
        const { searchTerm } = this.state;

        this.setState({ loading: true, error: null });

        try {
            // Clear cache when starting a new search
            if (page === 1) {
                this.userService.clearCache();
            }
            
            // If no search term, get all users, otherwise search with the term
            const result = searchTerm.trim() 
                ? await this.userService.getActiveUsers(page, searchTerm.trim())
                : await this.userService.getActiveUsers(page);

            this.setState({
                searchResults: result.users,
                currentPage: result.pagination.currentPage,
                totalPages: result.pagination.totalPages,
                totalRecords: result.pagination.totalRecords,
                hasNextPage: result.pagination.hasNextPage,
                hasPreviousPage: result.pagination.hasPreviousPage,
                loading: false,
                hasSearched: true
            });

        } catch (error) {
            this.setState({
                loading: false,
                error: `Failed to load users: ${error}`,
                hasSearched: true
            });
        }
    };

    private onClear = async (): Promise<void> => {
        this.setState({
            searchTerm: '',
            error: null
        });
        
        // Reload the first 25 users when clearing search
        await this.loadInitialUsers();
    };

    private isUserSelected = (user: SystemUser): boolean => {
        return this.state.selectedUsers.some(selectedUser => 
            (user.systemuserid && user.systemuserid === selectedUser.systemuserid) ||
            (user.internalemailaddress === selectedUser.internalemailaddress)
        );
    };

    private onUserCheckboxChange = (user: SystemUser, checked: boolean): void => {
        const { selectedUsers } = this.state;
        
        let newSelectedUsers: SystemUser[];
        
        if (checked) {
            if (!this.isUserSelected(user)) {
                newSelectedUsers = [...selectedUsers, user];
            } else {
                newSelectedUsers = selectedUsers;
            }
        } else {
            newSelectedUsers = selectedUsers.filter(selectedUser => 
                !(
                    (user.systemuserid && user.systemuserid === selectedUser.systemuserid) ||
                    (user.internalemailaddress === selectedUser.internalemailaddress)
                )
            );
        }
        
        this.setState({ selectedUsers: newSelectedUsers });
    };

    private onConfirmSelection = (): void => {
        this.props.onUsersSelected(this.state.selectedUsers);
    };

    private onClearSelection = (): void => {
        this.setState({ selectedUsers: [] });
    };

    private onSearchButtonClick = (): void => {
        this.onSearch(1);
    };

    private onKeyPress = (event: React.KeyboardEvent): void => {
        if (event.key === 'Enter') {
            this.onSearch(1);
        }
    };

    private onNextPage = async (): Promise<void> => {
        const { currentPage, hasNextPage, searchTerm } = this.state;
        if (!hasNextPage) return;

        this.setState({ loading: true, error: null });

        try {
            const result = await this.userService.getNextPage(currentPage, searchTerm.trim() || undefined);

            this.setState({
                searchResults: result.users,
                currentPage: result.pagination.currentPage,
                totalPages: result.pagination.totalPages,
                totalRecords: result.pagination.totalRecords,
                hasNextPage: result.pagination.hasNextPage,
                hasPreviousPage: result.pagination.hasPreviousPage,
                loading: false
            });

        } catch (error) {
            this.setState({
                loading: false,
                error: `Failed to load next page: ${error}`
            });
        }
    };

    private onPreviousPage = async (): Promise<void> => {
        const { hasPreviousPage, searchTerm } = this.state;
        if (!hasPreviousPage) return;

        this.setState({ loading: true, error: null });

        try {
            // For now, go back to first page (simplified approach)
            const result = await this.userService.getActiveUsers(1, searchTerm.trim() || undefined);

            this.setState({
                searchResults: result.users,
                currentPage: result.pagination.currentPage,
                totalPages: result.pagination.totalPages,
                totalRecords: result.pagination.totalRecords,
                hasNextPage: result.pagination.hasNextPage,
                hasPreviousPage: result.pagination.hasPreviousPage,
                loading: false
            });

        } catch (error) {
            this.setState({
                loading: false,
                error: `Failed to load previous page: ${error}`
            });
        }
    };

    public render(): React.ReactElement {
        const { onCancel } = this.props;
        const { searchTerm, searchResults, selectedUsers, loading, error, hasSearched, currentPage, totalPages, totalRecords, hasNextPage, hasPreviousPage } = this.state;

        return (
            <Stack tokens={{ childrenGap: 24 }} 
                   styles={{ 
                       root: { 
                           padding: '24px',
                           background: 'linear-gradient(145deg, #f8f9fa 0%, #ffffff 50%, #f1f3f5 100%)',
                           minHeight: '500px',
                           fontFamily: 'Segoe UI, "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif'
                       } 
                   }}>
                {/* Search Form */}
                <Stack tokens={{ childrenGap: 18 }} 
                       styles={{ 
                           root: { 
                               background: 'linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%)', 
                               padding: '24px', 
                               borderRadius: '12px', 
                               border: '2px solid #e9ecef',
                               boxShadow: '0 4px 16px rgba(0,0,0,0.08)',
                               position: 'relative',
                               overflow: 'hidden'
                           } 
                       }}>
                    <Stack tokens={{ childrenGap: 8 }}>
                        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                            Search Active AD Users
                        </Text>
                        <Text variant="small" styles={{ root: { color: '#605e5c', fontStyle: 'italic' } }}>
                            Showing only enabled users from Active Directory with assigned departments
                        </Text>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="end">
                        <Stack.Item grow>
                            <TextField
                                label="Search by Name"
                                placeholder="Enter first name, last name, or email..."
                                value={searchTerm}
                                onChange={(_, newValue) => this.setState({ searchTerm: newValue || '' })}
                                onKeyPress={this.onKeyPress}
                                disabled={loading}
                                styles={{
                                    root: { width: '100%' },
                                    fieldGroup: { borderRadius: '4px', border: '2px solid #e1e5e9' },
                                    field: { padding: '8px 12px' }
                                }}
                            />
                        </Stack.Item>
                        <PrimaryButton
                            text="Search"
                            iconProps={{ iconName: 'Search' }}
                            onClick={this.onSearchButtonClick}
                            disabled={loading || !searchTerm.trim()}
                            styles={{
                                root: { minWidth: '100px', height: '32px', borderRadius: '4px' },
                                rootHovered: { backgroundColor: '#106ebe' }
                            }}
                        />
                        <DefaultButton
                            text="Show All"
                            iconProps={{ iconName: 'People' }}
                            onClick={this.loadInitialUsers}
                            disabled={loading}
                            styles={{
                                root: { minWidth: '90px', height: '32px', borderRadius: '4px' }
                            }}
                        />
                        <DefaultButton
                            text="Clear"
                            iconProps={{ iconName: 'Clear' }}
                            onClick={this.onClear}
                            disabled={loading}
                            styles={{
                                root: { minWidth: '80px', height: '32px', borderRadius: '4px' }
                            }}
                        />
                    </Stack>
                </Stack>

                <Separator />

                {/* Error Display */}
                {error && (
                    <MessageBar messageBarType={MessageBarType.error}>
                        {error}
                    </MessageBar>
                )}

                {/* Selected Users Info */}
                {selectedUsers.length > 0 && (
                    <MessageBar messageBarType={MessageBarType.info}>
                        {selectedUsers.length} user(s) selected
                    </MessageBar>
                )}

                {/* Search Results */}
                {hasSearched && !loading && searchResults.length === 0 && !error && (
                    <MessageBar messageBarType={MessageBarType.info}>
                        No users found. Try adjusting your search criteria.
                    </MessageBar>
                )}

                {hasSearched && !loading && searchResults.length > 0 && (
                    <Stack tokens={{ childrenGap: 16 }}
                           styles={{ 
                               root: { 
                                   background: 'linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%)', 
                                   borderRadius: '12px', 
                                   border: '2px solid #e9ecef',
                                   boxShadow: '0 6px 20px rgba(0,0,0,0.1)',
                                   overflow: 'hidden'
                               } 
                           }}>
                        
                        {/* Results Header */}
                        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" 
                               styles={{ root: { padding: '16px 20px 0', borderBottom: '1px solid #f3f2f1' } }}>
                            <Text variant="medium" styles={{ root: { fontWeight: 600, color: '#2c3e50' } }}>
                                📋 Active AD Users (A-Z)
                            </Text>
                            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                                <Text variant="small" styles={{ root: { color: '#0078d4', fontWeight: 600 } }}>
                                    {searchResults.length} users • Page {currentPage} of {totalPages}
                                </Text>
                                <Text variant="small" styles={{ root: { color: '#6c757d', fontWeight: 500 } }}>
                                    • Sorted A-Z by first name
                                </Text>
                            </Stack>
                        </Stack>

                        {/* Results Grid */}
                        <div style={{ 
                            maxHeight: '400px', 
                            overflowY: 'auto',
                            margin: '0 20px',
                            border: '1px solid #edebe9',
                            borderRadius: '4px'
                        }}>
                            <div style={{ 
                                display: 'grid', 
                                gridTemplateColumns: '40px 1fr 1fr 2fr', 
                                gap: '0',
                                background: '#f8f9fa',
                                padding: '12px 16px',
                                fontWeight: 600,
                                fontSize: '14px',
                                color: '#323130',
                                borderBottom: '2px solid #e1e5e9'
                            }}>
                                <div></div>
                                <div>First Name</div>
                                <div>Last Name</div>
                                <div>Email</div>
                            </div>
                            {searchResults.map((user, index) => (
                                <div key={user.systemuserid || user.internalemailaddress || index} 
                                     style={{ 
                                        display: 'grid', 
                                        gridTemplateColumns: '40px 1fr 1fr 2fr', 
                                        gap: '0',
                                        padding: '12px 16px',
                                        borderBottom: '1px solid #f3f2f1',
                                        alignItems: 'center',
                                        background: this.isUserSelected(user) ? '#f0f6ff' : 'white',
                                        cursor: 'pointer'
                                     }}
                                     onClick={() => this.onUserCheckboxChange(user, !this.isUserSelected(user))}>
                                    <input
                                        type="checkbox"
                                        checked={this.isUserSelected(user)}
                                        onChange={(e) => {
                                            e.stopPropagation();
                                            this.onUserCheckboxChange(user, e.target.checked);
                                        }}
                                        style={{ 
                                            cursor: 'pointer',
                                            transform: 'scale(1.2)'
                                        }}
                                    />
                                    <div style={{ 
                                        fontSize: '14px', 
                                        color: '#323130',
                                        padding: '4px 0'
                                    }}>
                                        {user.firstname || 'N/A'}
                                    </div>
                                    <div style={{ 
                                        fontSize: '14px', 
                                        color: '#323130',
                                        padding: '4px 0'
                                    }}>
                                        {user.lastname || 'N/A'}
                                    </div>
                                    <div style={{ 
                                        fontSize: '14px', 
                                        color: '#605e5c',
                                        padding: '4px 0'
                                    }}>
                                        {user.internalemailaddress || 'No email'}
                                    </div>
                                </div>
                            ))}
                        </div>

                        {/* Pagination - Always show when we have results */}
                        {searchResults.length > 0 && (
                            <Stack horizontal horizontalAlign="space-between" verticalAlign="center"
                                   styles={{ 
                                       root: { 
                                           padding: '16px 20px', 
                                           borderTop: '1px solid #f3f2f1',
                                           backgroundColor: '#f8f9fa',
                                           borderRadius: '0 0 12px 12px'
                                       } 
                                   }}>
                                <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                                    <DefaultButton
                                        text="⬅ Previous"
                                        onClick={this.onPreviousPage}
                                        disabled={currentPage <= 1 || loading}
                                        styles={{ 
                                            root: { 
                                                minWidth: '110px',
                                                height: '36px',
                                                borderRadius: '6px',
                                                fontWeight: 500
                                            } 
                                        }}
                                    />
                                    <DefaultButton
                                        text="Next ➡"
                                        onClick={this.onNextPage}
                                        disabled={currentPage >= 20 || loading} // 495/25 = 20 pages max
                                        styles={{ 
                                            root: { 
                                                minWidth: '110px',
                                                height: '36px',
                                                borderRadius: '6px',
                                                fontWeight: 500,
                                                backgroundColor: hasNextPage ? '#0078d4' : undefined,
                                                color: hasNextPage ? 'white' : undefined,
                                                border: hasNextPage ? 'none' : undefined
                                            } 
                                        }}
                                    />
                                </Stack>
                                <Stack tokens={{ childrenGap: 4 }}>
                                    <Text variant="small" styles={{ root: { color: '#605e5c', fontWeight: 600 } }}>
                                        Page {currentPage}
                                    </Text>
                                    <Text variant="small" styles={{ root: { color: '#6c757d', fontSize: '12px' } }}>
                                        Showing {searchResults.length} users
                                    </Text>
                                </Stack>
                            </Stack>
                        )}
                    </Stack>
                )}

                {/* Loading Indicator */}
                {loading && (
                    <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}
                           styles={{ root: { padding: '40px', background: 'white', borderRadius: '8px', border: '1px solid #e1e5e9' } }}>
                        <div style={{ 
                            width: '40px', 
                            height: '40px', 
                            border: '4px solid #f3f3f3', 
                            borderTop: '4px solid #0078d4', 
                            borderRadius: '50%', 
                            animation: 'spin 1s linear infinite' 
                        }}></div>
                        <Text variant="medium" styles={{ root: { color: '#605e5c' } }}>
                            Searching users...
                        </Text>
                    </Stack>
                )}

                <Separator styles={{ root: { margin: '20px 0' } }} />

                {/* Action Buttons */}
                <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 15 }}
                       styles={{ root: { padding: '20px', background: 'white', borderRadius: '8px', border: '1px solid #e1e5e9' } }}>
                    <DefaultButton
                        text="Clear Selection"
                        iconProps={{ iconName: 'Clear' }}
                        onClick={this.onClearSelection}
                        disabled={selectedUsers.length === 0 || loading}
                        styles={{
                            root: { 
                                minWidth: '120px', 
                                height: '36px',
                                borderRadius: '4px',
                                border: '1px solid #d1d1d1'
                            }
                        }}
                    />
                    <Stack horizontal tokens={{ childrenGap: 12 }}>
                        <DefaultButton
                            text="Cancel"
                            onClick={onCancel}
                            disabled={loading}
                            styles={{
                                root: { 
                                    minWidth: '100px', 
                                    height: '36px',
                                    borderRadius: '4px',
                                    border: '1px solid #d1d1d1'
                                }
                            }}
                        />
                        <PrimaryButton
                            text={selectedUsers.length > 0 ? `Select ${selectedUsers.length} User(s)` : 'Select Users'}
                            onClick={this.onConfirmSelection}
                            disabled={selectedUsers.length === 0 || loading}
                            styles={{
                                root: { 
                                    minWidth: '140px', 
                                    height: '36px',
                                    borderRadius: '4px',
                                    backgroundColor: '#0078d4',
                                    border: 'none'
                                },
                                rootHovered: {
                                    backgroundColor: '#106ebe'
                                }
                            }}
                        />
                    </Stack>
                </Stack>
            </Stack>
        );
    }
}