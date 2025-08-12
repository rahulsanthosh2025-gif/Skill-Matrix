export interface ISkillLog {
  EmployeeEmail: string;
  Skills: string;
  ExpectedProficiency: string;
  ActualProficiency: string;
  ActionPlan: string;
  Timeline: string;
  SkillGap: string;
  Comments: string;
  Attachments?: string;
}

export interface ISkillMatrixProps {
  skillLogs: ISkillLog[];
  isLoading: boolean;
}

export interface ISkillMatrixState {
  skillLogs: ISkillLog[];
  isLoading: boolean;
  error?: string;
}