import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISkillLog } from '../models/ISkillLog';

export class SharePointService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getSkillLogs(): Promise<ISkillLog[]> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('SkillLogs')/items?$select=EmployeeEmail,Skills,ExpectedProficiency,ActualProficiency,ActionPlan,Timeline,SkillGap,Comments,Attachments`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((item: any) => ({
          EmployeeEmail: item.EmployeeEmail || '',
          Skills: item.Skills || '',
          ExpectedProficiency: item.ExpectedProficiency || '',
          ActualProficiency: item.ActualProficiency || '',
          ActionPlan: item.ActionPlan || '',
          Timeline: item.Timeline || '',
          SkillGap: item.SkillGap || '',
          Comments: item.Comments || '',
          Attachments: item.Attachments || ''
        }));
      } else {
        throw new Error(`Failed to fetch skill logs: ${response.statusText}`);
      }
    } catch (error) {
      console.error('Error fetching skill logs:', error);
      throw error;
    }
  }

  public async addSkillLog(skillLog: ISkillLog): Promise<void> {
    try {
      const body = JSON.stringify({
        EmployeeEmail: skillLog.EmployeeEmail,
        Skills: skillLog.Skills,
        ExpectedProficiency: skillLog.ExpectedProficiency,
        ActualProficiency: skillLog.ActualProficiency,
        ActionPlan: skillLog.ActionPlan,
        Timeline: skillLog.Timeline,
        SkillGap: skillLog.SkillGap,
        Comments: skillLog.Comments,
        Attachments: skillLog.Attachments
      });

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('SkillLogs')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: body
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to add skill log: ${response.statusText}`);
      }
    } catch (error) {
      console.error('Error adding skill log:', error);
      throw error;
    }
  }
}