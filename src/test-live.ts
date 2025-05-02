import { DataverseService, DataverseConfig } from './DataverseService.js';
import { config } from 'dotenv';

// Load environment variables from .env file
config();

async function testLiveEnvironment() {
  // Validate required environment variables
  const requiredEnvVars = [
    'DATAVERSE_URL',
    'DATAVERSE_CLIENT_ID',
    'DATAVERSE_CLIENT_SECRET',
    'DATAVERSE_TENANT_ID'
  ] as const;

  for (const envVar of requiredEnvVars) {
    if (!process.env[envVar]) {
      console.error(`Error: ${envVar} is not set in environment variables`);
      console.error('Please create a .env file based on .env.template');
      process.exit(1);
    }
  }

  const config: DataverseConfig = {
    organizationUrl: process.env.DATAVERSE_URL!,
    clientId: process.env.DATAVERSE_CLIENT_ID!,
    clientSecret: process.env.DATAVERSE_CLIENT_SECRET!,
    tenantId: process.env.DATAVERSE_TENANT_ID!
  };

  const service = new DataverseService(config);

  try {
    // Create a new account
    console.log('Creating new account...');
    const newAccount = await service.createRecord('accounts', {
      name: 'Hitachi Solutions',
      description: 'Global consulting and IT services company',
      websiteurl: 'https://www.hitachi-solutions.com',
      industrycode: 1 // Consulting
    });
    
    console.log('Created account:', newAccount);

    // Update the account
    console.log('\nUpdating account...');
    const updatedAccount = await service.updateRecord('accounts', newAccount.accountid, {
      description: 'Updated by Dataverse MCP test at ' + new Date().toISOString(),
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