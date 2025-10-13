/**
 * Service for handling SystemUser table operations with pagination and search
 */

export interface SystemUser {
    systemuserid: string;
    firstname: string;
    lastname: string;
    internalemailaddress: string;
    isdisabled?: boolean;
    domainname?: string;
    title?: string;
}

export interface PaginationInfo {
    currentPage: number;
    pageSize: number;
    totalRecords: number;
    totalPages: number;
    hasNextPage: boolean;
    hasPreviousPage: boolean;
}

export interface SystemUserSearchResult {
    users: SystemUser[];
    pagination: PaginationInfo;
}

export class SystemUserService {
    private context: ComponentFramework.Context<any>;
    private pageSize: number = 5;
    private lastUserIdPerPage: Map<number, string> = new Map(); // Store last user ID for each page
    private pageData: Map<number, SystemUser[]> = new Map(); // Cache page data

    constructor(context: ComponentFramework.Context<any>) {
        this.context = context;
    }

    /**
     * Retrieves active system users with pagination and optional search
     * @param page Current page number (1-based)
     * @param searchTerm Optional search term for firstname/lastname fuzzy matching
     * @param nextPageUrl Optional URL for next page (used internally)
     * @returns Promise with users and pagination info
     */
    public async getActiveUsers(page: number = 1, searchTerm?: string): Promise<SystemUserSearchResult> {
        try {
            console.log(`SystemUserService: Getting users for page ${page} with search: "${searchTerm || 'none'}"`);
            
            // Check if we have cached data for this exact page and search term
            const cacheKey = `${page}_${searchTerm || 'all'}`;
            
            // Build cursor-based query
            const query = this.buildODataQueryWithCursor(searchTerm, page);
            console.log(`SystemUserService: Query: ${query}`);
                
            // Execute the query
            const result = await this.context.webAPI.retrieveMultipleRecords(
                "systemuser",
                query
            );

            // Map results to SystemUser interface
            const users: SystemUser[] = result.entities.map((entity: any) => ({
                systemuserid: entity.systemuserid,
                firstname: entity.firstname || '',
                lastname: entity.lastname || '',
                internalemailaddress: entity.internalemailaddress || '',
                isdisabled: entity.isdisabled || false,
                domainname: entity.domainname || '',
                title: entity.title || ''
            }));

            // Store the last user ID for cursor-based pagination
            if (users.length > 0) {
                this.lastUserIdPerPage.set(page, users[users.length - 1].systemuserid);
                console.log(`SystemUserService: Stored last user ID for page ${page}: ${users[users.length - 1].systemuserid}`);
            }

            // Calculate pagination info
            const totalRecords = 495; // Known total from your environment
            const totalPages = Math.ceil(totalRecords / this.pageSize); // 495/5 = 99 pages
            
            const pagination: PaginationInfo = {
                currentPage: page,
                pageSize: this.pageSize,
                totalRecords: totalRecords,
                totalPages: totalPages,
                hasNextPage: page < totalPages,
                hasPreviousPage: page > 1
            };

            // Debug logging
            console.log('SystemUserService - A-Z Sorted pagination result:', {
                usersCount: users.length,
                currentPage: page,
                hasNextPage: pagination.hasNextPage,
                hasPreviousPage: pagination.hasPreviousPage,
                firstUser: users.length > 0 ? `${users[0].firstname} ${users[0].lastname}` : 'none',
                lastUser: users.length > 0 ? `${users[users.length - 1].firstname} ${users[users.length - 1].lastname}` : 'none',
                sortOrder: 'First Name A-Z, then Last Name A-Z'
            });

            return {
                users,
                pagination
            };

        } catch (error: any) {
            console.error('Error retrieving system users:', error);
            
            // Better error handling to get more specific error details
            let errorMessage = 'Unknown error occurred';
            if (error && typeof error === 'object') {
                if (error.message) {
                    errorMessage = error.message;
                } else if (error.errorCode) {
                    errorMessage = `Error Code: ${error.errorCode} - ${error.errorMessage || 'API Error'}`;
                } else {
                    errorMessage = JSON.stringify(error);
                }
            } else if (typeof error === 'string') {
                errorMessage = error;
            }
            
            throw new Error(`Failed to retrieve system users: ${errorMessage}`);
        }
    }

    /**
     * Execute a custom query URL (for pagination) using Web API context
     */
    private async executeCustomQuery(url: string): Promise<any> {
        try {
            // Extract the query part from the full URL
            const urlObj = new URL(url);
            const queryParams = urlObj.search;
            
            // Use the context's webAPI to execute the query
            const result = await this.context.webAPI.retrieveMultipleRecords(
                "systemuser",
                queryParams
            );
            
            return result;
        } catch (error) {
            console.error('Error executing custom query:', error);
            throw error;
        }
    }

    /**
     * Gets next page of results using cached next page URL or new query
     * @param currentPage Current page number
     * @param searchTerm Search term used
     * @returns Promise with next page results
     */
    public async getNextPage(currentPage: number, searchTerm?: string): Promise<SystemUserSearchResult> {
        console.log(`SystemUserService: Getting next page ${currentPage + 1}`);
        return this.getActiveUsers(currentPage + 1, searchTerm);
    }



    /**
     * Gets previous page
     * @param currentPage Current page number
     * @param searchTerm Search term used
     * @returns Promise with previous page results
     */
    public async getPreviousPage(currentPage: number, searchTerm?: string): Promise<SystemUserSearchResult> {
        const previousPage = Math.max(1, currentPage - 1);
        console.log(`SystemUserService: Getting previous page ${previousPage}`);
        return this.getActiveUsers(previousPage, searchTerm);
    }

    /**
     * Clears the pagination cache
     */
    public clearCache(): void {
        this.lastUserIdPerPage.clear();
        this.pageData.clear();
    }

    /**
     * Sets the page size for pagination
     * @param size Number of records per page
     */
    public setPageSize(size: number): void {
        this.pageSize = size;
        this.clearCache(); // Clear cache when page size changes
    }

    /**
     * Builds OData query string with cursor-based pagination for active AD users with departments
     * @param searchTerm Optional search term
     * @param page Page number for cursor-based filtering
     * @returns OData query string
     */
    private buildODataQueryWithCursor(searchTerm?: string, page: number = 1): string {
        let query = "?$select=systemuserid,firstname,lastname,internalemailaddress,isdisabled,domainname,title";
        
        let filters: string[] = [];

        // Filter for active users only (not disabled)
        filters.push("isdisabled eq false");
        
        // Filter for AD users only (domainname should not be null/empty)
        filters.push("domainname ne null");
        
        // Filter for users with a department/title (indicating they have a role)
        filters.push("title ne null");

        // Add search filter if provided
        if (searchTerm && searchTerm.trim()) {
            const searchFilter = this.buildSearchFilter(searchTerm.trim());
            filters.push(searchFilter);
        }

        // Add cursor-based pagination filter
        if (page > 1) {
            // Get the last user ID from the previous page for cursor-based pagination
            const previousPageLastId = this.lastUserIdPerPage.get(page - 1);
            if (previousPageLastId) {
                // Use systemuserid for cursor since it's unique and part of our ordering
                filters.push(`systemuserid gt '${previousPageLastId}'`);
            }
        }

        // Combine filters
        if (filters.length > 0) {
            query += `&$filter=${filters.join(' and ')}`;
        }

        // Add ordering - A-Z by first name first, then last name for alphabetical sorting
        query += "&$orderby=firstname asc,lastname asc,systemuserid asc";

        // Add pagination - only use $top, no $skip
        query += `&$top=${this.pageSize}`;

        return query;
    }

    /**
     * Legacy method for backward compatibility
     */
    private buildODataQuery(searchTerm?: string): string {
        return this.buildODataQueryWithCursor(searchTerm, 1);
    }

    /**
     * Builds search filter for fuzzy matching on firstname, lastname, and email
     * @param searchTerm Search term
     * @returns OData filter string
     */
    private buildSearchFilter(searchTerm: string): string {
        // Use startswith function which is more widely supported
        // Escape single quotes in search term
        const escapedSearchTerm = searchTerm.replace(/'/g, "''");
        
        // Create filters for multiple fields using startswith for fuzzy matching
        const firstnameFilter = `startswith(firstname,'${escapedSearchTerm}')`;
        const lastnameFilter = `startswith(lastname,'${escapedSearchTerm}')`;
        const emailFilter = `startswith(internalemailaddress,'${escapedSearchTerm}')`;
        
        // Also check if the search term matches when split by space (for "John Doe" type searches)
        const searchWords = searchTerm.trim().split(/\s+/);
        let combinedFilters: string[] = [firstnameFilter, lastnameFilter, emailFilter];
        
        if (searchWords.length > 1) {
            // For multi-word searches, check if firstname startswith first word AND lastname startswith second word
            const firstWord = searchWords[0].replace(/'/g, "''");
            const secondWord = searchWords[1].replace(/'/g, "''");
            const fullNameFilter = `(startswith(firstname,'${firstWord}') and startswith(lastname,'${secondWord}'))`;
            combinedFilters.push(fullNameFilter);
            
            // Also try reverse order (lastname firstname)
            const reverseNameFilter = `(startswith(lastname,'${firstWord}') and startswith(firstname,'${secondWord}'))`;
            combinedFilters.push(reverseNameFilter);
        }
        
        // Combine all filters with OR logic
        return `(${combinedFilters.join(' or ')})`;
    }

    /**
     * Gets total count of records matching the search criteria
     * @param searchTerm Optional search term
     * @returns Promise with total count
     */
    private async getTotalCount(searchTerm?: string): Promise<number> {
        try {
            let query = "?$select=systemuserid&$count=true";
            
            // Remove restrictive filters that might cause permission issues
            let filters: string[] = [];

            // Add search filter if provided
            if (searchTerm && searchTerm.trim()) {
                const searchFilter = this.buildSearchFilter(searchTerm.trim());
                filters.push(searchFilter);
            }

            // Combine filters
            if (filters.length > 0) {
                query += `&$filter=${filters.join(' and ')}`;
            }

            const result = await this.context.webAPI.retrieveMultipleRecords(
                "systemuser",
                query
            );

            return (result as any)["@odata.count"] || result.entities.length;

        } catch (error) {
            console.error('Error getting total count:', error);
            return 0;
        }
    }

    /**
     * Searches for a specific user by ID
     * @param userId System User ID
     * @returns Promise with user details or null if not found
     */
    public async getUserById(userId: string): Promise<SystemUser | null> {
        try {
            const query = "?$select=systemuserid,firstname,lastname,internalemailaddress,isdisabled,domainname,title";
            
            const result = await this.context.webAPI.retrieveRecord(
                "systemuser",
                userId,
                query
            );

            return {
                systemuserid: result.systemuserid,
                firstname: result.firstname || '',
                lastname: result.lastname || '',
                internalemailaddress: result.internalemailaddress || '',
                isdisabled: result.isdisabled || false,
                domainname: result.domainname || '',
                title: result.title || ''
            };

        } catch (error) {
            console.error('Error retrieving user by ID:', error);
            return null;
        }
    }

    /**
     * Gets multiple users by their IDs
     * @param userIds Array of system user IDs
     * @returns Promise with array of users
     */
    public async getUsersByIds(userIds: string[]): Promise<SystemUser[]> {
        try {
            if (!userIds || userIds.length === 0) {
                return [];
            }

            // Build filter for multiple IDs
            const idFilters = userIds.map(id => `systemuserid eq '${id}'`).join(' or ');
            const query = `?$select=systemuserid,firstname,lastname,internalemailaddress,isdisabled,domainname,title&$filter=${idFilters}`;

            const result = await this.context.webAPI.retrieveMultipleRecords(
                "systemuser",
                query
            );

            return result.entities.map(entity => ({
                systemuserid: entity.systemuserid,
                firstname: entity.firstname || '',
                lastname: entity.lastname || '',
                internalemailaddress: entity.internalemailaddress || '',
                isdisabled: entity.isdisabled || false,
                domainname: entity.domainname || '',
                title: entity.title || ''
            }));

        } catch (error) {
            console.error('Error retrieving users by IDs:', error);
            return [];
        }
    }
}