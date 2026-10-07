/* global module */

/**
 * @typedef {Object} CategoryDefinition
 * @property {string} label Display name; detections are grouped by this value.
 * @property {number} weight Points per detection.
 * @property {boolean} [highSensitivity] Counts towards the escalation bonus.
 */

/**
 * @typedef {Object} CategoryDetection
 * @property {string} label
 * @property {number} count
 */

/**
 * @typedef {Object} RiskLevelDefinition
 * @property {string} name
 * @property {number} min Inclusive lower bound.
 * @property {number} max Inclusive upper bound.
 */

/**
 * @typedef {Object} EscalationRule
 * @property {number} minHighSensitivityCategories
 * @property {number} bonus
 */

/**
 * @typedef {Object} ScoringConfig
 * @property {RiskLevelDefinition[]} riskLevels Ordered by ascending threshold.
 * @property {EscalationRule} escalation
 */

/**
 * @typedef {Object} CategoryScoreBreakdown
 * @property {string} category
 * @property {number} detections
 * @property {number} weight
 * @property {number} categoryScore
 * @property {boolean} highSensitivity
 */

/**
 * @typedef {Object} RiskScoreResult
 * @property {CategoryScoreBreakdown[]} breakdown
 * @property {number} baseScore
 * @property {number} highSensitivityCategoryCount
 * @property {number} escalationBonus
 * @property {number} finalScore
 * @property {string} riskLevel
 */

/** @type {ScoringConfig} */
const DEFAULT_SCORING_CONFIG = {
  riskLevels: [
    { name: "Low", min: 1, max: 5 },
    { name: "Medium", min: 6, max: 15 },
    { name: "High", min: 16, max: 25 },
    { name: "Very High", min: 26, max: 40 },
    { name: "Critical", min: 41, max: Number.POSITIVE_INFINITY },
  ],
  escalation: { minHighSensitivityCategories: 2, bonus: 10 },
};

/**
 * @param {number} score
 * @param {RiskLevelDefinition[]} [levels]
 * @returns {string} Empty when the score is below the first level (i.e. nothing detected).
 */
function getRiskLevel(score, levels = DEFAULT_SCORING_CONFIG.riskLevels) {
  if (score < levels[0].min) {
    return "";
  }

  const level = levels.find((candidate) => score >= candidate.min && score <= candidate.max);
  return level ? level.name : levels[levels.length - 1].name;
}

/**
 * Category-agnostic: a new category only needs a CategoryDefinition, not a change here.
 * @param {CategoryDefinition[]} definitions
 * @param {CategoryDetection[]} detections
 * @param {ScoringConfig} [config]
 * @returns {RiskScoreResult}
 */
function calculateRiskScore(definitions, detections, config = DEFAULT_SCORING_CONFIG) {
  const definitionByLabel = new Map(
    definitions.map((definition) => [definition.label, definition])
  );
  const countsByLabel = new Map();

  detections.forEach(({ label, count }) => {
    if (definitionByLabel.has(label) && count > 0) {
      countsByLabel.set(label, (countsByLabel.get(label) || 0) + count);
    }
  });

  /** @type {CategoryScoreBreakdown[]} */
  const breakdown = [];
  countsByLabel.forEach((count, label) => {
    const definition = definitionByLabel.get(label);
    breakdown.push({
      category: label,
      detections: count,
      weight: definition.weight,
      categoryScore: definition.weight * count,
      highSensitivity: Boolean(definition.highSensitivity),
    });
  });

  breakdown.sort((a, b) => b.categoryScore - a.categoryScore);

  const baseScore = breakdown.reduce((total, row) => total + row.categoryScore, 0);
  const highSensitivityCategoryCount = breakdown.filter((row) => row.highSensitivity).length;
  const escalationBonus =
    highSensitivityCategoryCount >= config.escalation.minHighSensitivityCategories
      ? config.escalation.bonus
      : 0;
  const finalScore = baseScore + escalationBonus;

  return {
    breakdown,
    baseScore,
    highSensitivityCategoryCount,
    escalationBonus,
    finalScore,
    riskLevel: getRiskLevel(finalScore, config.riskLevels),
  };
}

module.exports = { DEFAULT_SCORING_CONFIG, getRiskLevel, calculateRiskScore };
