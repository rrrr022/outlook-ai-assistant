import { Router, Request, Response } from 'express';
import { AutomationRule, RuleCondition, RuleAction } from '../../shared/types';
import { v4 as uuidv4 } from 'uuid';

const router = Router();

// In-memory storage for rules (in production, use a database)
let automationRules: AutomationRule[] = [];

/**
 * GET /api/automation/rules
 * Get all automation rules
 */
router.get('/rules', (req: Request, res: Response) => {
  res.json({
    success: true,
    rules: automationRules,
  });
});

/**
 * POST /api/automation/rules
 * Create a new automation rule
 */
router.post('/rules', (req: Request, res: Response) => {
  try {
    const { name, description, conditions, conditionLogic, actions, isEnabled } = req.body;

    if (!name) {
      return res.status(400).json({
        success: false,
        error: 'Rule name is required',
      });
    }

    const newRule: AutomationRule = {
      id: uuidv4(),
      name,
      description,
      isEnabled: isEnabled ?? true,
      priority: automationRules.length + 1,
      conditions: conditions || [],
      conditionLogic: conditionLogic || 'and',
      actions: actions || [],
      createdAt: new Date(),
      updatedAt: new Date(),
      triggerCount: 0,
    };

    automationRules.push(newRule);

    res.status(201).json({
      success: true,
      rule: newRule,
    });
  } catch (error: any) {
    console.error('Create rule error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to create rule',
    });
  }
});

/**
 * PUT /api/automation/rules/:id
 * Update an automation rule
 */
router.put('/rules/:id', (req: Request, res: Response) => {
  try {
    const { id } = req.params;
    const updates = req.body;

    const ruleIndex = automationRules.findIndex(r => r.id === id);
    if (ruleIndex === -1) {
      return res.status(404).json({
        success: false,
        error: 'Rule not found',
      });
    }

    automationRules[ruleIndex] = {
      ...automationRules[ruleIndex],
      ...updates,
      updatedAt: new Date(),
    };

    res.json({
      success: true,
      rule: automationRules[ruleIndex],
    });
  } catch (error: any) {
    console.error('Update rule error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to update rule',
    });
  }
});

/**
 * DELETE /api/automation/rules/:id
 * Delete an automation rule
 */
router.delete('/rules/:id', (req: Request, res: Response) => {
  try {
    const { id } = req.params;

    const ruleIndex = automationRules.findIndex(r => r.id === id);
    if (ruleIndex === -1) {
      return res.status(404).json({
        success: false,
        error: 'Rule not found',
      });
    }

    automationRules.splice(ruleIndex, 1);

    res.json({
      success: true,
      message: 'Rule deleted successfully',
    });
  } catch (error: any) {
    console.error('Delete rule error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to delete rule',
    });
  }
});

/**
 * POST /api/automation/evaluate
 * Evaluate rules against an email/event
 */
router.post('/evaluate', (req: Request, res: Response) => {
  try {
    const { email, event } = req.body;
    const matchingRules: AutomationRule[] = [];
    const suggestedActions: RuleAction[] = [];

    for (const rule of automationRules) {
      if (!rule.isEnabled) continue;

      const conditionsMet = evaluateConditions(rule.conditions, rule.conditionLogic, { email, event });
      
      if (conditionsMet) {
        matchingRules.push(rule);
        suggestedActions.push(...rule.actions);
        
        // Update trigger count
        rule.triggerCount++;
        rule.lastTriggered = new Date();
      }
    }

    res.json({
      success: true,
      matchingRules,
      suggestedActions,
    });
  } catch (error: any) {
    console.error('Evaluate rules error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to evaluate rules',
    });
  }
});

/**
 * Helper function to evaluate conditions
 */
function evaluateConditions(
  conditions: RuleCondition[],
  logic: 'and' | 'or',
  context: { email?: any; event?: any }
): boolean {
  if (conditions.length === 0) return false;

  const results = conditions.map(condition => evaluateCondition(condition, context));

  if (logic === 'and') {
    return results.every(r => r);
  } else {
    return results.some(r => r);
  }
}

/**
 * Helper function to evaluate a single condition
 */
function evaluateCondition(condition: RuleCondition, context: { email?: any; event?: any }): boolean {
  const { type, operator, value } = condition;
  const { email, event } = context;

  let targetValue: any;

  // Get the target value based on condition type
  switch (type) {
    case 'sender':
      targetValue = email?.sender || email?.senderEmail;
      break;
    case 'subject':
      targetValue = email?.subject;
      break;
    case 'body':
      targetValue = email?.body || email?.preview;
      break;
    case 'hasAttachment':
      targetValue = email?.hasAttachments;
      break;
    case 'importance':
      targetValue = email?.importance;
      break;
    case 'isRead':
      targetValue = email?.isRead;
      break;
    default:
      return false;
  }

  if (targetValue === undefined) return false;

  // Evaluate based on operator
  switch (operator) {
    case 'equals':
      return targetValue === value;
    case 'contains':
      return String(targetValue).toLowerCase().includes(String(value).toLowerCase());
    case 'startsWith':
      return String(targetValue).toLowerCase().startsWith(String(value).toLowerCase());
    case 'endsWith':
      return String(targetValue).toLowerCase().endsWith(String(value).toLowerCase());
    case 'greaterThan':
      return Number(targetValue) > Number(value);
    case 'lessThan':
      return Number(targetValue) < Number(value);
    default:
      return false;
  }
}

export default router;
