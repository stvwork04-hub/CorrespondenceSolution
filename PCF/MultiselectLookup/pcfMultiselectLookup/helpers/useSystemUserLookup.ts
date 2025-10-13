import * as React from 'react';
import { SystemUserService, SystemUser, SystemUserSearchResult } from './SystemUserService';

export interface UseSystemUserLookupOptions {
    context: ComponentFramework.Context<any>;
    pageSize?: number;
    autoLoad?: boolean;
}

export interface UseSystemUserLookupReturn {
    users: SystemUser[];
    loading: boolean;
    error: string | null;
    pagination: {
        currentPage: number;
        totalPages: number;
        totalRecords: number;
        hasNextPage: boolean;
        hasPreviousPage: boolean;
    };
    actions: {
        loadUsers: (page?: number, searchTerm?: string) => Promise<void>;
        nextPage: () => Promise<void>;
        previousPage: () => Promise<void>;
        goToPage: (page: number) => Promise<void>;
        search: (searchTerm: string) => Promise<void>;
        getUserById: (userId: string) => Promise<SystemUser | null>;
        getUsersByIds: (userIds: string[]) => Promise<SystemUser[]>;
    };
}

/**
 * Custom hook for managing system user lookup with pagination and search
 */
export function useSystemUserLookup(options: UseSystemUserLookupOptions): UseSystemUserLookupReturn {
    const { context, pageSize = 25, autoLoad = true } = options;
    
    const [users, setUsers] = React.useState<SystemUser[]>([]);
    const [loading, setLoading] = React.useState<boolean>(false);
    const [error, setError] = React.useState<string | null>(null);
    const [currentPage, setCurrentPage] = React.useState<number>(1);
    const [totalPages, setTotalPages] = React.useState<number>(0);
    const [totalRecords, setTotalRecords] = React.useState<number>(0);
    const [hasNextPage, setHasNextPage] = React.useState<boolean>(false);
    const [hasPreviousPage, setHasPreviousPage] = React.useState<boolean>(false);
    const [currentSearchTerm, setCurrentSearchTerm] = React.useState<string>('');
    
    const userService = React.useMemo(() => {
        const service = new SystemUserService(context);
        if (pageSize !== 25) {
            service.setPageSize(pageSize);
        }
        return service;
    }, [context, pageSize]);

    const loadUsers = React.useCallback(async (page: number = 1, searchTerm?: string): Promise<void> => {
        setLoading(true);
        setError(null);

        try {
            const result: SystemUserSearchResult = await userService.getActiveUsers(page, searchTerm);
            
            setUsers(result.users);
            setCurrentPage(result.pagination.currentPage);
            setTotalPages(result.pagination.totalPages);
            setTotalRecords(result.pagination.totalRecords);
            setHasNextPage(result.pagination.hasNextPage);
            setHasPreviousPage(result.pagination.hasPreviousPage);
            
            if (searchTerm !== undefined) {
                setCurrentSearchTerm(searchTerm);
            }

        } catch (err) {
            const errorMessage = `Failed to load users: ${err}`;
            setError(errorMessage);
            setUsers([]);
        } finally {
            setLoading(false);
        }
    }, [userService]);

    const nextPage = React.useCallback(async (): Promise<void> => {
        if (hasNextPage) {
            await loadUsers(currentPage + 1, currentSearchTerm);
        }
    }, [hasNextPage, currentPage, currentSearchTerm, loadUsers]);

    const previousPage = React.useCallback(async (): Promise<void> => {
        if (hasPreviousPage) {
            await loadUsers(currentPage - 1, currentSearchTerm);
        }
    }, [hasPreviousPage, currentPage, currentSearchTerm, loadUsers]);

    const goToPage = React.useCallback(async (page: number): Promise<void> => {
        if (page >= 1 && page <= totalPages) {
            await loadUsers(page, currentSearchTerm);
        }
    }, [totalPages, currentSearchTerm, loadUsers]);

    const search = React.useCallback(async (searchTerm: string): Promise<void> => {
        await loadUsers(1, searchTerm);
    }, [loadUsers]);

    const getUserById = React.useCallback(async (userId: string): Promise<SystemUser | null> => {
        try {
            return await userService.getUserById(userId);
        } catch (err) {
            console.error('Error getting user by ID:', err);
            return null;
        }
    }, [userService]);

    const getUsersByIds = React.useCallback(async (userIds: string[]): Promise<SystemUser[]> => {
        try {
            return await userService.getUsersByIds(userIds);
        } catch (err) {
            console.error('Error getting users by IDs:', err);
            return [];
        }
    }, [userService]);

    // Auto-load on mount if enabled
    React.useEffect(() => {
        if (autoLoad) {
            loadUsers();
        }
    }, [autoLoad, loadUsers]);

    return {
        users,
        loading,
        error,
        pagination: {
            currentPage,
            totalPages,
            totalRecords,
            hasNextPage,
            hasPreviousPage
        },
        actions: {
            loadUsers,
            nextPage,
            previousPage,
            goToPage,
            search,
            getUserById,
            getUsersByIds
        }
    };
}