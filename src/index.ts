#!/usr/bin/env node

import { config } from 'dotenv';
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from 'zod';
import { DataverseService, DataverseConfig } from './DataverseService.js';

// Load environment variables from .env file
config();

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

// Environment configuration
const DATAVERSE_CONFIG: DataverseConfig = {
  organizationUrl: process.env.DATAVERSE_URL!,
  clientId: process.env.DATAVERSE_CLIENT_ID!,
  clientSecret: process.env.DATAVERSE_CLIENT_SECRET!,
  tenantId: process.env.DATAVERSE_TENANT_ID!,
};

// Create server instance
const server = new McpServer({
  name: "dataverse-mcp",
  version: "1.0.0",
});

let dataverseService: DataverseService | null = null;

// Function to initialize DataverseService on demand
function getDataverseService(): DataverseService {
  if (!dataverseService) {
    dataverseService = new DataverseService(DATAVERSE_CONFIG);
  }
  return dataverseService;
}

// Pre-defined Dataverse Prompts
const dataversePrompts = {
  ENTITY_OVERVIEW: (entityName: string) => 
    `## Dataverse Entity: ${entityName}\n\n` +
    `This is an overview of the '${entityName}' entity in Microsoft Dataverse:\n\n` +
    `### Entity Details\n{{entity_details}}\n\n` +
    `### Attributes\n{{key_attributes}}\n\n` +
    `### Relationships\n{{relationships}}\n\n` +
    `You can query this entity using OData filters against the plural name.`,

  ATTRIBUTE_DETAILS: (entityName: string, attributeName: string) =>
    `## Attribute: ${attributeName}\n\n` +
    `Details for the '${attributeName}' attribute of the '${entityName}' entity:\n\n` +
    `{{attribute_details}}\n\n` +
    `### Usage Notes\n` +
    `- Data Type: {{data_type}}\n` +
    `- Required: {{required}}\n` +
    `- Max Length: {{max_length}}`,

  QUERY_TEMPLATE: (entityNamePlural: string) =>
    `## OData Query Template for ${entityNamePlural}\n\n` +
    `Use this template to build queries against the ${entityNamePlural} entity:\n\n` +
    `\`\`\`\n${entityNamePlural}?$select={{selected_fields}}&$filter={{filter_conditions}}&$orderby={{order_by}}&$top={{max_records}}\n\`\`\`\n\n` +
    `### Common Filter Examples\n` +
    `- Equals: \`name eq 'Contoso'\`\n` +
    `- Contains: \`contains(name, 'Contoso')\`\n` +
    `- Greater than date: \`createdon gt 2023-01-01T00:00:00Z\`\n` +
    `- Multiple conditions: \`name eq 'Contoso' and statecode eq 0\``,

  RELATIONSHIP_MAP: (entityName: string) =>
    `## Relationship Map for ${entityName}\n\n` +
    `This shows all relationships for the '${entityName}' entity:\n\n` +
    `### One-to-Many Relationships (${entityName} as Primary)\n{{one_to_many_primary}}\n\n` +
    `### One-to-Many Relationships (${entityName} as Related)\n{{one_to_many_related}}\n\n` +
    `### Many-to-Many Relationships\n{{many_to_many}}\n\n`
};

// Register prompts with the server using the correct method signature
server.prompt(
  "entity-overview", 
  "Get an overview of a Dataverse entity",
  {
    entityName: z.string().describe("The logical name of the entity")
  },
  async (args) => {
    try {
      const service = getDataverseService();
      const entityName = args.entityName;
      
      const [metadata, attributes] = await Promise.all([
        service.getEntityMetadata(entityName),
        service.getEntityAttributes(entityName)
      ]);
      
      const entityDetails = `- Display Name: ${metadata.DisplayName?.UserLocalizedLabel?.Label || entityName}\n` +
        `- Schema Name: ${metadata.SchemaName}\n` +
        `- Description: ${metadata.Description?.UserLocalizedLabel?.Label || 'No description'}\n` +
        `- Primary Key: ${metadata.PrimaryIdAttribute}\n` +
        `- Primary Name: ${metadata.PrimaryNameAttribute}`;
        
      const keyAttributes = attributes.value
        .map((attr: any) => {
          const attrType = attr["@odata.type"] || attr.odata?.type || "Unknown type";
          return `- ${attr.LogicalName}: ${attrType}`;
        })
        .join('\n');
        
      const relationships = await service.getEntityRelationships(entityName);
      const oneToManyCount = relationships.oneToMany.value.length;
      const manyToManyCount = relationships.manyToMany.value.length;
      
      const relationshipsSummary = `- One-to-Many Relationships: ${oneToManyCount}\n` +
                                  `- Many-to-Many Relationships: ${manyToManyCount}`;
      
      let promptContent = dataversePrompts.ENTITY_OVERVIEW(entityName);
      promptContent = promptContent
        .replace('{{entity_details}}', entityDetails)
        .replace('{{key_attributes}}', keyAttributes)
        .replace('{{relationships}}', relationshipsSummary);
      
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: promptContent
            }
          }
        ]
      };
    } catch (error: any) {
      console.error(`Error handling entity-overview prompt:`, error);
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: `Error: ${error.message}`
            }
          }
        ]
      };
    }
  }
);

server.prompt(
  "attribute-details",
  "Get detailed information about a specific entity attribute/field",
  {
    entityName: z.string().describe("The logical name of the entity"),
    attributeName: z.string().describe("The logical name of the attribute"),
  },
  async (args) => {
    try {
      const service = getDataverseService();
      const { entityName, attributeName } = args;
      
      const attribute = await service.getEntityAttribute(entityName, attributeName);
      
      const attrDetails = `- Display Name: ${attribute.DisplayName?.UserLocalizedLabel?.Label || attributeName}\n` +
        `- Description: ${attribute.Description?.UserLocalizedLabel?.Label || 'No description'}\n` +
        `- Type: ${attribute.AttributeType}\n` +
        `- Format: ${attribute.Format || 'N/A'}\n` +
        `- Is Required: ${attribute.RequiredLevel?.Value || 'No'}\n` +
        `- Is Searchable: ${attribute.IsValidForAdvancedFind || false}`;
        
      let promptContent = dataversePrompts.ATTRIBUTE_DETAILS(entityName, attributeName);
      promptContent = promptContent
        .replace('{{attribute_details}}', attrDetails)
        .replace('{{data_type}}', attribute.AttributeType)
        .replace('{{required}}', attribute.RequiredLevel?.Value || 'No')
        .replace('{{max_length}}', attribute.MaxLength || 'N/A');
      
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: promptContent
            }
          }
        ]
      };
    } catch (error: any) {
      console.error(`Error handling attribute-details prompt:`, error);
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: `Error: ${error.message}`
            }
          }
        ]
      };
    }
  }
);

server.prompt(
  "query-template",
  "Get a template for querying a Dataverse entity",
  {
    entityName: z.string().describe("The logical name of the entity"),
  },
  async (args) => {
    try {
      const service = getDataverseService();
      const entityName = args.entityName;
      
      const metadata = await service.getEntityMetadata(entityName);
      const entityNamePlural = metadata.EntitySetName;
      
      const attributes = await service.getEntityAttributes(entityName);
      const selectFields = attributes.value
        .filter((attr: any) => attr.IsValidForRead === true && !attr.AttributeOf)
        .slice(0, 5)
        .map((attr: any) => attr.LogicalName)
        .join(',');
        
      let promptContent = dataversePrompts.QUERY_TEMPLATE(entityNamePlural);
      promptContent = promptContent
        .replace('{{selected_fields}}', selectFields)
        .replace('{{filter_conditions}}', `${metadata.PrimaryNameAttribute} eq 'Example'`)
        .replace('{{order_by}}', `${metadata.PrimaryNameAttribute} asc`)
        .replace('{{max_records}}', '50');
      
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: promptContent
            }
          }
        ]
      };
    } catch (error: any) {
      console.error(`Error handling query-template prompt:`, error);
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: `Error: ${error.message}`
            }
          }
        ]
      };
    }
  }
);

server.prompt(
  "relationship-map",
  "Get a list of relationships for a Dataverse entity",
  {
    entityName: z.string().describe("The logical name of the entity"),
  },
  async (args) => {
    try {
      const service = getDataverseService();
      const entityName = args.entityName;
      
      const relationships = await service.getEntityRelationships(entityName);
      
      const oneToManyPrimary = relationships.oneToMany.value
        .filter((rel: any) => rel.ReferencingEntity !== entityName)
        .map((rel: any) => `- ${rel.SchemaName}: ${entityName} (1) → ${rel.ReferencingEntity} (N)`)
        .join('\n');
        
      const oneToManyRelated = relationships.oneToMany.value
        .filter((rel: any) => rel.ReferencingEntity === entityName)
        .map((rel: any) => `- ${rel.SchemaName}: ${rel.ReferencedEntity} (1) → ${entityName} (N)`)
        .join('\n');
        
      const manyToMany = relationships.manyToMany.value
        .map((rel: any) => {
          const otherEntity = rel.Entity1LogicalName === entityName ? rel.Entity2LogicalName : rel.Entity1LogicalName;
          return `- ${rel.SchemaName}: ${entityName} (N) ↔ ${otherEntity} (N)`;
        })
        .join('\n');
      
      let promptContent = dataversePrompts.RELATIONSHIP_MAP(entityName);
      promptContent = promptContent
        .replace('{{one_to_many_primary}}', oneToManyPrimary || 'None found')
        .replace('{{one_to_many_related}}', oneToManyRelated || 'None found')
        .replace('{{many_to_many}}', manyToMany || 'None found');
      
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: promptContent
            }
          }
        ]
      };
    } catch (error: any) {
      console.error(`Error handling relationship-map prompt:`, error);
      return {
        messages: [
          {
            role: "assistant",
            content: {
              type: "text",
              text: `Error: ${error.message}`
            }
          }
        ]
      };
    }
  }
);

server.tool(
  "get-entity-metadata",
  "Get metadata about a Dataverse entity",
  {
    entityName: z.string().describe("The logical name of the entity"),
  },
  async ({ entityName }) => {
    try {
      const service = getDataverseService();
      const metadata = await service.getEntityMetadata(entityName);
      
      const metadataStr = JSON.stringify(metadata, null, 2);
      
      return {
        content: [
          {
            type: "text",
            text: `Entity metadata for '${entityName}':\n\n${metadataStr}`,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error getting entity metadata:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to get entity metadata: ${error.message}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "get-entity-attributes",
  "Get attributes/fields of a Dataverse entity",
  {
    entityName: z.string().describe("The logical name of the entity"),
  },
  async ({ entityName }) => {
    try {
      const service = getDataverseService();
      const attributes = await service.getEntityAttributes(entityName);
      
      const attributesStr = JSON.stringify(attributes, null, 2);
      
      return {
        content: [
          {
            type: "text",
            text: `Attributes for entity '${entityName}':\n\n${attributesStr}`,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error getting entity attributes:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to get entity attributes: ${error.message}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "get-entity-attribute",
  "Get a specific attribute/field of a Dataverse entity",
  {
    entityName: z.string().describe("The logical name of the entity"),
    attributeName: z.string().describe("The logical name of the attribute")
  },
  async ({ entityName, attributeName }) => {
    try {
      const service = getDataverseService();
      const attribute = await service.getEntityAttribute(entityName, attributeName);
      
      const attributeStr = JSON.stringify(attribute, null, 2);
      
      return {
        content: [
          {
            type: "text",
            text: `Attribute '${attributeName}' for entity '${entityName}':\n\n${attributeStr}`,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error getting entity attribute:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to get entity attribute: ${error.message}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "get-entity-relationships",
  "Get relationships (one-to-many and many-to-many) for a Dataverse entity",
  {
    entityName: z.string().describe("The logical name of the entity"),
  },
  async ({ entityName }) => {
    try {
      const service = getDataverseService();
      const relationships = await service.getEntityRelationships(entityName);
      
      const relationshipsStr = JSON.stringify(relationships, null, 2);
      
      return {
        content: [
          {
            type: "text",
            text: `Relationships for entity '${entityName}':\n\n${relationshipsStr}`,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error getting entity relationships:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to get entity relationships: ${error.message}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "get-global-option-set",
  "Get a global option set definition by name",
  {
    optionSetName: z.string().describe("The name of the global option set"),
  },
  async ({ optionSetName }) => {
    try {
      const service = getDataverseService();
      const optionSet = await service.getGlobalOptionSet(optionSetName);
      
      const optionSetStr = JSON.stringify(optionSet, null, 2);
      
      return {
        content: [
          {
            type: "text",
            text: `Global option set '${optionSetName}':\n\n${optionSetStr}`,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error getting global option set:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to get global option set: ${error.message}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "get-record",
  "Get a specific record by entity name (plural) and ID",
  {
    entityNamePlural: z.string().describe("The plural name of the entity (e.g., 'accounts', 'contacts')"),
    recordId: z.string().describe("The GUID of the record"),
  },
  async ({ entityNamePlural, recordId }) => {
    try {
      const service = getDataverseService();
      const record = await service.getRecord(entityNamePlural, recordId);
      
      const recordStr = JSON.stringify(record, null, 2);
      
      return {
        content: [
          {
            type: "text",
            text: `Record from '${entityNamePlural}' with ID '${recordId}':\n\n${recordStr}`,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error getting record:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to get record: ${error.message}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "query-records",
  "Query records using an OData filter expression",
  {
    entityNamePlural: z.string().describe("The plural name of the entity (e.g., 'accounts', 'contacts')"),
    filter: z.string().describe("OData filter expression (e.g., \"name eq 'test'\" or \"createdon gt 2023-01-01\")"),
    maxRecords: z.number().optional().describe("Maximum number of records to retrieve (default: 50)"),
  },
  async ({ entityNamePlural, filter, maxRecords }) => {
    try {
      const service = getDataverseService();
      const records = await service.queryRecords(entityNamePlural, filter, maxRecords || 50);
      
      const recordsStr = JSON.stringify(records, null, 2);
      const recordCount = records.value?.length || 0;
      
      return {
        content: [
          {
            type: "text",
            text: `Retrieved ${recordCount} records from '${entityNamePlural}' with filter '${filter}':\n\n${recordsStr}`,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error querying records:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to query records: ${error.message}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "use-dataverse-prompt",
  "Use a predefined prompt template for Dataverse entities",
  {
    promptType: z.enum([
      "ENTITY_OVERVIEW", 
      "ATTRIBUTE_DETAILS", 
      "QUERY_TEMPLATE", 
      "RELATIONSHIP_MAP"
    ]).describe("The type of prompt template to use"),
    entityName: z.string().describe("The logical name of the entity"),
    attributeName: z.string().optional().describe("The logical name of the attribute (required for ATTRIBUTE_DETAILS prompt)"),
  },
  async ({ promptType, entityName, attributeName }) => {
    try {
      const service = getDataverseService();
      
      let promptContent = "";
      let replacements: Record<string, string> = {};
      
      switch (promptType) {
        case "ENTITY_OVERVIEW": {
          const [metadata, attributes] = await Promise.all([
            service.getEntityMetadata(entityName),
            service.getEntityAttributes(entityName)
          ]);
          
          const entityDetails = `- Display Name: ${metadata.DisplayName?.UserLocalizedLabel?.Label || entityName}\n` +
            `- Schema Name: ${metadata.SchemaName}\n` +
            `- Description: ${metadata.Description?.UserLocalizedLabel?.Label || 'No description'}\n` +
            `- Primary Key: ${metadata.PrimaryIdAttribute}\n` +
            `- Primary Name: ${metadata.PrimaryNameAttribute}`;
            
          const keyAttributes = attributes.value
            .map((attr: any) => {
                const attrType = attr["@odata.type"] || attr.odata?.type || "Unknown type";
                return `- ${attr.LogicalName}: ${attrType}`;
              })
            .join('\n');
            
          const relationships = await service.getEntityRelationships(entityName);
          const oneToManyCount = relationships.oneToMany.value.length;
          const manyToManyCount = relationships.manyToMany.value.length;
          
          const relationshipsSummary = `- One-to-Many Relationships: ${oneToManyCount}\n` +
                                      `- Many-to-Many Relationships: ${manyToManyCount}`;
          
          promptContent = dataversePrompts.ENTITY_OVERVIEW(entityName);
          replacements = {
            '{{entity_details}}': entityDetails,
            '{{key_attributes}}': keyAttributes,
            '{{relationships}}': relationshipsSummary
          };
          break;
        }
        
        case "ATTRIBUTE_DETAILS": {
          if (!attributeName) {
            throw new Error("attributeName is required for ATTRIBUTE_DETAILS prompt");
          }
          
          const attribute = await service.getEntityAttribute(entityName, attributeName);
          
          const attrDetails = `- Display Name: ${attribute.DisplayName?.UserLocalizedLabel?.Label || attributeName}\n` +
            `- Description: ${attribute.Description?.UserLocalizedLabel?.Label || 'No description'}\n` +
            `- Type: ${attribute.AttributeType}\n` +
            `- Format: ${attribute.Format || 'N/A'}\n` +
            `- Is Required: ${attribute.RequiredLevel?.Value || 'No'}\n` +
            `- Is Searchable: ${attribute.IsValidForAdvancedFind || false}`;
            
          promptContent = dataversePrompts.ATTRIBUTE_DETAILS(entityName, attributeName);
          replacements = {
            '{{attribute_details}}': attrDetails,
            '{{data_type}}': attribute.AttributeType,
            '{{required}}': attribute.RequiredLevel?.Value || 'No',
            '{{max_length}}': attribute.MaxLength || 'N/A'
          };
          break;
        }
        
        case "QUERY_TEMPLATE": {
          const metadata = await service.getEntityMetadata(entityName);
          const entityNamePlural = metadata.EntitySetName;
          
          const attributes = await service.getEntityAttributes(entityName);
          const selectFields = attributes.value
            .slice(0, 5)
            .map((attr: any) => attr.LogicalName)
            .join(',');
            
          promptContent = dataversePrompts.QUERY_TEMPLATE(entityNamePlural);
          replacements = {
            '{{selected_fields}}': selectFields,
            '{{filter_conditions}}': `${metadata.PrimaryNameAttribute} eq 'Example'`,
            '{{order_by}}': `${metadata.PrimaryNameAttribute} asc`,
            '{{max_records}}': '50'
          };
          break;
        }
        
        case "RELATIONSHIP_MAP": {
          const relationships = await service.getEntityRelationships(entityName);
          
          const oneToManyPrimary = relationships.oneToMany.value
            .filter((rel: any) => rel.ReferencingEntity !== entityName)
            .map((rel: any) => `- ${rel.SchemaName}: ${entityName} (1) → ${rel.ReferencingEntity} (N)`)
            .join('\n');
            
          const oneToManyRelated = relationships.oneToMany.value
            .filter((rel: any) => rel.ReferencingEntity === entityName)
            .map((rel: any) => `- ${rel.SchemaName}: ${rel.ReferencedEntity} (1) → ${entityName} (N)`)
            .join('\n');
            
          const manyToMany = relationships.manyToMany.value
            .map((rel: any) => {
              const otherEntity = rel.Entity1LogicalName === entityName ? rel.Entity2LogicalName : rel.Entity1LogicalName;
              return `- ${rel.SchemaName}: ${entityName} (N) ↔ ${otherEntity} (N)`;
            })
            .join('\n');
          
          promptContent = dataversePrompts.RELATIONSHIP_MAP(entityName);
          replacements = {
            '{{one_to_many_primary}}': oneToManyPrimary || 'None found',
            '{{one_to_many_related}}': oneToManyRelated || 'None found',
            '{{many_to_many}}': manyToMany || 'None found'
          };
          break;
        }
      }
      
      for (const [placeholder, value] of Object.entries(replacements)) {
        promptContent = promptContent.replace(placeholder, value);
      }
      
      return {
        content: [
          {
            type: "text",
            text: promptContent,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error using Dataverse prompt:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to use Dataverse prompt: ${error.message}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "create-record",
  "Create a new record in a Dataverse entity",
  {
    entityNamePlural: z.string().describe("The plural name of the entity (e.g., 'accounts', 'contacts')"),
    data: z.object({}).passthrough().describe("The record data to create"),
  },
  async ({ entityNamePlural, data }) => {
    try {
      const service = getDataverseService();
      const record = await service.createRecord(entityNamePlural, data);
      
      return {
        content: [
          {
            type: "text",
            text: `Successfully created record in '${entityNamePlural}':\n\n${JSON.stringify(record, null, 2)}`,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error creating record:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to create record: ${error.message}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "update-record",
  "Update an existing record in a Dataverse entity",
  {
    entityNamePlural: z.string().describe("The plural name of the entity (e.g., 'accounts', 'contacts')"),
    recordId: z.string().describe("The GUID of the record to update"),
    data: z.object({}).passthrough().describe("The record data to update"),
  },
  async ({ entityNamePlural, recordId, data }) => {
    try {
      const service = getDataverseService();
      const record = await service.updateRecord(entityNamePlural, recordId, data);
      
      return {
        content: [
          {
            type: "text",
            text: `Successfully updated record in '${entityNamePlural}':\n\n${JSON.stringify(record, null, 2)}`,
          },
        ],
      };
    } catch (error: any) {
      console.error("Error updating record:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to update record: ${error.message}`,
          },
        ],
      };
    }
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Initializing Dataverse MCP Server...");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});