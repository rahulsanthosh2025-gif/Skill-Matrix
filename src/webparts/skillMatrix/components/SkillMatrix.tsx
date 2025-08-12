import * as React from 'react';
import { ISkillMatrixProps } from './ISkillMatrixProps';
import { ISkillLogItem } from '../../../../models/ISkillLogItem';
import styles from './SkillMatrix.module.scss';
import { sp } from '@pnp/sp/presets/all';

const LIST_NAME = 'SkillLogs';

const SkillMatrix: React.FC<ISkillMatrixProps> = () => {
  const [items, setItems] = React.useState<ISkillLogItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);

  const loadItems = async (): Promise<void> => {
    setLoading(true);
    try {
      const data = await sp.web.lists
        .getByTitle(LIST_NAME)
        .items.select(
          'Id',
          'EmployeeEmail',
          'Skills',
          'ExpectedProficiency',
          'ActualProficiency',
          'ActionPlan',
          'Timeline',
          'SkillGap',
          'Comments'
        )
        .get<ISkillLogItem[]>();

      setItems(data);
    } catch (error) {
      console.error(`[SkillMatrix] Error loading items:`, error);
    }
    setLoading(false);
  };

  React.useEffect(() => {
    loadItems();
  }, []);

  if (loading) {
    return <div className={styles.skillMatrix}>Loading...</div>;
  }

  return (
    <div className={styles.skillMatrix}>
      <h2 className={styles.header}>Skill Matrix</h2>
      <table className={styles.table}>
        <thead>
          <tr>
            <th>Employee Email</th>
            <th>Skills</th>
            <th>Expected</th>
            <th>Actual</th>
            <th>Action Plan</th>
            <th>Timeline</th>
            <th>Skill Gap</th>
            <th>Comments</th>
          </tr>
        </thead>
        <tbody>
          {items.map((item) => (
            <tr key={item.Id}>
              <td>{item.EmployeeEmail}</td>
              <td>{item.Skills}</td>
              <td>{item.ExpectedProficiency}</td>
              <td>{item.ActualProficiency}</td>
              <td>{item.ActionPlan}</td>
              <td>{item.Timeline}</td>
              <td>{item.SkillGap}</td>
              <td>{item.Comments}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default SkillMatrix;