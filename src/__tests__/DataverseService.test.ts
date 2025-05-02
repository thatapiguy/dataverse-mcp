import { DataverseService, DataverseConfig } from '../DataverseService.js';
import axios, { AxiosInstance } from 'axios';

// Mock axios
jest.mock('axios');
const mockedAxios = jest.mocked(axios);

// Mock MSAL
jest.mock('@azure/msal-node', () => ({
  ConfidentialClientApplication: jest.fn().mockImplementation(() => ({
    acquireTokenByClientCredential: jest.fn().mockResolvedValue({
      accessToken: 'mock-token',
      expiresOn: new Date(Date.now() + 3600000) // 1 hour from now
    })
  }))
}));

describe('DataverseService', () => {
  let service: DataverseService;
  const mockConfig: DataverseConfig = {
    organizationUrl: 'https://test.crm.dynamics.com',
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    tenantId: 'test-tenant-id'
  };

  beforeEach(() => {
    service = new DataverseService(mockConfig);
    jest.clearAllMocks();
  });

  describe('createRecord', () => {
    it('should successfully create a record', async () => {
      const mockData = { name: 'Test Account' };
      const mockResponse = {
        data: { id: '12345', ...mockData }
      };

      mockedAxios.mockResolvedValueOnce(mockResponse);

      const result = await service.createRecord('accounts', mockData);

      expect(mockedAxios).toHaveBeenCalledWith({
        method: 'POST',
        url: 'https://test.crm.dynamics.com/api/data/v9.2/accounts',
        headers: {
          'Authorization': 'Bearer mock-token',
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          'Prefer': 'return=representation'
        },
        data: mockData
      });

      expect(result).toEqual(mockResponse.data);
    });

    it('should handle create record errors', async () => {
      const mockError = new Error('Create failed');
      mockedAxios.mockRejectedValueOnce(mockError);

      await expect(service.createRecord('accounts', {}))
        .rejects
        .toThrow('Failed to create record: Error: Create failed');
    });
  });

  describe('updateRecord', () => {
    it('should successfully update a record', async () => {
      const mockData = { name: 'Updated Account' };
      const mockRecordId = '12345';
      const mockResponse = {
        data: { id: mockRecordId, ...mockData }
      };

      mockedAxios.mockResolvedValueOnce(mockResponse);

      const result = await service.updateRecord('accounts', mockRecordId, mockData);

      expect(mockedAxios).toHaveBeenCalledWith({
        method: 'PATCH',
        url: 'https://test.crm.dynamics.com/api/data/v9.2/accounts(12345)',
        headers: {
          'Authorization': 'Bearer mock-token',
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          'If-Match': '*',
          'Prefer': 'return=representation'
        },
        data: mockData
      });

      expect(result).toEqual(mockResponse.data);
    });

    it('should handle update record errors', async () => {
      const mockError = new Error('Update failed');
      mockedAxios.mockRejectedValueOnce(mockError);

      await expect(service.updateRecord('accounts', '12345', {}))
        .rejects
        .toThrow('Failed to update record: Error: Update failed');
    });
  });

  describe('queryRecordsWithPagination', () => {
    it('should handle pagination correctly', async () => {
      const mockData = {
        value: [{ id: '1' }, { id: '2' }],
        '@odata.nextLink': 'https://test.crm.dynamics.com/api/data/v9.2/accounts?$skiptoken=123'
      };

      mockedAxios.mockResolvedValueOnce({ data: mockData });

      const result = await service.queryRecordsWithPagination('accounts', "statecode eq 0", 50);

      expect(mockedAxios).toHaveBeenCalledWith({
        method: 'GET',
        url: 'https://test.crm.dynamics.com/api/data/v9.2/accounts?$filter=statecode%20eq%200&$top=50',
        headers: {
          'Authorization': 'Bearer mock-token',
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0'
        }
      });

      expect(result).toEqual({
        value: mockData.value,
        nextPageUrl: mockData['@odata.nextLink']
      });
    });

    it('should use continuation token when provided', async () => {
      const mockData = {
        value: [{ id: '3' }, { id: '4' }]
      };

      const continuationToken = 'https://test.crm.dynamics.com/api/data/v9.2/accounts?$skiptoken=123';
      mockedAxios.mockResolvedValueOnce({ data: mockData });

      const result = await service.queryRecordsWithPagination('accounts', "statecode eq 0", 50, continuationToken);

      expect(mockedAxios).toHaveBeenCalledWith({
        method: 'GET',
        url: continuationToken,
        headers: {
          'Authorization': 'Bearer mock-token',
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0'
        }
      });

      expect(result).toEqual({
        value: mockData.value,
        nextPageUrl: undefined
      });
    });

    it('should handle errors', async () => {
      const mockError = new Error('Query failed');
      mockedAxios.mockRejectedValueOnce(mockError);

      await expect(service.queryRecordsWithPagination('accounts', "statecode eq 0"))
        .rejects
        .toThrow('Failed to query records: Error: Query failed');
    });
  });
});