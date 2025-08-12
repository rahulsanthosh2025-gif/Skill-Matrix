import * as React from 'react';
import { ISkillLog, ISkillMatrixProps, ISkillMatrixState } from '../../../models/ISkillLog';
import { SharePointService } from '../../../services/SharePointService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import styles from './SkillMatrix.module.scss';

export interface ISkillMatrixComponentProps extends ISkillMatrixProps {
  context: WebPartContext;
}

export class SkillMatrix extends React.Component<ISkillMatrixComponentProps, ISkillMatrixState> {
  private sharePointService: SharePointService;

  constructor(props: ISkillMatrixComponentProps) {
    super(props);
    
    this.state = {
      skillLogs: [],
      isLoading: true,
      error: undefined
    };

    this.sharePointService = new SharePointService(this.props.context);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadSkillLogs();
  }

  private async loadSkillLogs(): Promise<void> {
    try {
      this.setState({ isLoading: true, error: undefined });
      const skillLogs = await this.sharePointService.getSkillLogs();
      this.setState({ skillLogs, isLoading: false });
    } catch (error) {
      console.error('Error loading skill logs:', error);
      this.setState({ 
        isLoading: false, 
        error: `Failed to load skill logs: ${error.message}` 
      });
    }
  }

  public render(): React.ReactElement<ISkillMatrixComponentProps> {
    const { skillLogs, isLoading, error } = this.state;

    if (isLoading) {
      return (
        <div className={styles.skillMatrix}>
          <div className={styles.loadingContainer}>
            <Spinner label="Loading skill matrix data..." />
          </div>
        </div>
      );
    }

    if (error) {
      return (
        <div className={styles.skillMatrix}>
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        </div>
      );
    }

    return (
      <div className={styles.skillMatrix}>
        <div className={styles.header}>
          <h2 className={styles.title}>SKILL MATRIX</h2>
        </div>
        
        <div className={styles.tableContainer}>
          <table className={styles.skillTable}>
            <thead>
              <tr>
                <th>Skills</th>
                <th>Expected Proficiency</th>
                <th>Actual Proficiency</th>
                <th>Action Plan</th>
                <th>Timeline</th>
                <th>Skill Gap</th>
                <th>Employee Comments</th>
                <th>Attach Document</th>
              </tr>
            </thead>
            <tbody>
              {skillLogs.length === 0 ? (
                <tr>
                  <td colSpan={8} className={styles.noData}>
                    No skill logs found. Please add some data to the SkillLogs list.
                  </td>
                </tr>
              ) : (
                skillLogs.map((log, index) => (
                  <tr key={index}>
                    <td>{log.Skills}</td>
                    <td>{log.ExpectedProficiency}</td>
                    <td>{log.ActualProficiency}</td>
                    <td>{log.ActionPlan}</td>
                    <td>{log.Timeline}</td>
                    <td>{log.SkillGap}</td>
                    <td>{log.Comments}</td>
                    <td>
                      {log.Attachments && (
                        <a href={log.Attachments} target="_blank" rel="noopener noreferrer">
                          📎 View
                        </a>
                      )}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        <div className={styles.refreshContainer}>
          <button 
            className={styles.refreshButton}
            onClick={() => this.loadSkillLogs()}
            disabled={isLoading}
          >
            🔄 Refresh Data
          </button>
        </div>
      </div>
    );
  }
}