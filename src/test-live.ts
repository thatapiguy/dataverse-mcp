import { PowerPlatformService, PowerPlatformConfig } from './PowerPlatformService.js';

async function testLiveEnvironment() {
  const config: PowerPlatformConfig = {
    organizationUrl: process.env.POWERPLATFORM_URL || '',
    clientId: process.env.POWERPLATFORM_CLIENT_ID || '',
    clientSecret: process.env.POWERPLATFORM_CLIENT_SECRET || '',
    tenantId: process.env.POWERPLATFORM_TENANT_ID || ''
  };

  const service = new PowerPlatformService(config);

  try {
    // Create a new account
    console.log('Creating new account...');
    const newAccount = await service.createRecord('accounts', {
      name: 'Test Account ' + new Date().toISOString(),
      telephone1: '555-0123',
      emailaddress1: 'test@example.com',
      description: 'Created by PowerPlatform MCP test'
    });
    
    console.log('Created account:', newAccount);

    // Update the account
    console.log('\nUpdating account...');
    const updatedAccount = await service.updateRecord('accounts', newAccount.accountid, {
      description: 'Updated by PowerPlatform MCP test at ' + new Date().toISOString(),
      telephone1: '555-9999'
    });

    console.log('Updated account:', updatedAccount);

    // Query to verify
    console.log('\nVerifying with query...');
    const queryResult = await service.queryRecords(
      'accounts',
      `accountid eq ${newAccount.accountid}`,
      1
    );

    console.log('Query result:', queryResult);

  } catch (error) {
    console.error('Error during testing:', error);
  }
}

testLiveEnvironment().catch(console.error);