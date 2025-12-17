"""AI-powered analyzer for redlines based on legal playbook."""

from typing import List, Dict, Optional
import os
from openai import OpenAI
from anthropic import Anthropic
from playbook_loader import PlaybookLoader


class AIAnalyzer:
    """Analyzes redlines using AI models against a legal playbook."""
    
    def __init__(
        self,
        playbook: PlaybookLoader,
        provider: str = "openai",
        model: str = "gpt-4"
    ):
        """Initialize analyzer with playbook and AI configuration."""
        self.playbook = playbook
        self.provider = provider.lower()
        self.model = model
        
        if self.provider == "openai":
            api_key = os.getenv('OPENAI_API_KEY')
            if not api_key:
                raise ValueError("OPENAI_API_KEY not found in environment variables")
            self.client = OpenAI(api_key=api_key)
        elif self.provider == "anthropic":
            api_key = os.getenv('ANTHROPIC_API_KEY')
            if not api_key:
                raise ValueError("ANTHROPIC_API_KEY not found in environment variables")
            self.client = Anthropic(api_key=api_key)
        else:
            raise ValueError(f"Unsupported provider: {provider}. Use 'openai' or 'anthropic'")
    
    def analyze_redlines(
        self,
        redlines: List[Dict],
        document_text: str,
        context: Optional[str] = None
    ) -> List[Dict]:
        """Analyze EACH redline individually against the playbook using AI.
        
        CRITICAL: Each redline is analyzed separately to ensure individual attention
        and to provide specific guidance for each tracked change.
        """
        if not redlines:
            return []
        
        print(f"Analyzing {len(redlines)} redline(s) individually...")
        all_analyses = []
        
        # Analyze EACH redline individually
        for idx, redline in enumerate(redlines, 1):
            redline_type = redline.get('type', 'Unknown')
            if redline_type == 'replacement':
                old_text = redline.get('old_text', '')
                new_text = redline.get('new_text', '')
                print(f"  Analyzing redline {idx} of {len(redlines)}: {redline_type} - '{old_text}' → '{new_text}'")
            else:
                print(f"  Analyzing redline {idx} of {len(redlines)}: {redline_type} - {redline.get('text', '')[:50]}...")
            
            # Format this single redline for analysis
            redlines_summary = self._format_single_redline_for_analysis(redline, idx)
            playbook_text = self.playbook.get_playbook_text()
            
            # Build prompt for this individual redline
            prompt = self._build_analysis_prompt(
                playbook_text,
                redlines_summary,
                document_text,
                context
            )
            
            # Get AI response for this redline
            try:
                analysis = self._call_ai(prompt)
                print(f"    AI response received ({len(analysis)} chars)")
            except Exception as e:
                print(f"    ✗ ERROR calling AI: {type(e).__name__}: {e}")
                all_analyses.append({
                    'redline': redline,
                    'playbook_principle': '',
                    'assessment': f'AI analysis failed: {str(e)}',
                    'response': 'Please review this change manually',
                    'fallbacks': '',
                    'risk_level': 'Medium',
                    'comment_text': 'Please review this change against the legal playbook.',
                    'auto_redline_action': 'comment_only',
                    'auto_redline_text': ''
                })
                continue
            
            # Parse AI response and add the redline reference
            try:
                parsed = self._parse_ai_response(analysis, [redline], redline_number=1)
                if parsed and len(parsed) > 0:
                    # Ensure the redline is included in the analysis
                    parsed[0]['redline'] = redline
                    all_analyses.extend(parsed)
                    print(f"    ✓ Analysis parsed successfully")
                else:
                    print(f"    ⚠ Parsing returned empty result, creating fallback analysis")
                    # If parsing failed, create a basic analysis entry
                    all_analyses.append({
                        'redline': redline,
                        'playbook_principle': '',
                        'assessment': 'Analysis parsing failed - AI response format may be incorrect',
                        'response': 'Please review this change manually',
                        'fallbacks': '',
                        'risk_level': 'Medium',
                        'comment_text': 'Please review this change against the legal playbook.',
                        'auto_redline_action': 'comment_only',
                        'auto_redline_text': ''
                    })
            except Exception as e:
                print(f"    ✗ ERROR parsing AI response: {type(e).__name__}: {e}")
                import traceback
                print(f"    Traceback: {traceback.format_exc()}")
                # If parsing failed, create a basic analysis entry
                all_analyses.append({
                    'redline': redline,
                    'playbook_principle': '',
                    'assessment': f'Analysis parsing error: {str(e)}',
                    'response': 'Please review this change manually',
                    'fallbacks': '',
                    'risk_level': 'Medium',
                    'comment_text': 'Please review this change against the legal playbook.',
                    'auto_redline_action': 'comment_only',
                    'auto_redline_text': ''
                })
        
        print(f"✓ Completed analysis of {len(all_analyses)} redline(s)")
        return all_analyses
    
    def _format_single_redline_for_analysis(self, redline: Dict, number: int) -> str:
        """Format a single redline for AI analysis."""
        redline_type = redline.get('type', 'Unknown')
        text_info = ""
        
        # Handle replacements properly - show both old and new text
        if redline_type == 'replacement':
            old_text = redline.get('old_text', '')
            new_text = redline.get('new_text', '')
            text_info = f"  Old Text (deleted): {old_text}\n  New Text (inserted): {new_text}"
        else:
            text_info = f"  Text: {redline.get('text', '')}"
        
        return (
            f"Redline #{number}:\n"
            f"  Type: {redline_type}\n"
            f"{text_info}\n"
            f"  Author: {redline.get('author', 'Unknown')}\n"
            f"  Date: {redline.get('date', 'Unknown')}\n"
        )
    
    def _format_redlines_for_analysis(self, redlines: List[Dict]) -> str:
        """Format redlines for AI analysis."""
        formatted = []
        for idx, redline in enumerate(redlines, 1):
            redline_type = redline.get('type', 'Unknown')
            text_info = ""
            
            if redline_type == 'replacement':
                old_text = redline.get('old_text', '')
                new_text = redline.get('new_text', '')
                text_info = f"  Old Text (deleted): {old_text}\n  New Text (inserted): {new_text}"
            else:
                text_info = f"  Text: {redline.get('text', '')}"
            
            formatted.append(
                f"Redline #{idx}:\n"
                f"  Type: {redline_type}\n"
                f"{text_info}\n"
                f"  Author: {redline.get('author', 'Unknown')}\n"
                f"  Date: {redline.get('date', 'Unknown')}\n"
            )
        return '\n'.join(formatted)
    
    def _build_analysis_prompt(
        self,
        playbook_text: str,
        redlines_summary: str,
        document_text: str,
        context: Optional[str]
    ) -> str:
        """Build the prompt for AI analysis."""
        prompt = f"""You are a legal technology assistant analyzing redlines (tracked changes) in a legal document against a legal playbook.

LEGAL PLAYBOOK:
{playbook_text}

DOCUMENT CONTEXT:
{document_text[:2000]}...

REDLINES TO ANALYZE:
{redlines_summary}

TASK:
Analyze each redline and determine if it should be accepted, rejected with a counter-redline, or just commented on. For each redline:

1. FIRST: Find the specific playbook clause/principle that applies to this redline
2. QUOTE the exact playbook text in "playbook_principle" - this is REQUIRED
3. THEN assess whether the change aligns with or violates the playbook
4. Determine the appropriate AUTO-REDLINE ACTION:
   - "accept": The change is acceptable per the playbook (no counter-redline needed)
   - "reject_restore": Reject the change and restore the original text (for deletions: restore deleted text; for replacements: restore old text)
   - "reject_replace": Reject the change and insert specific alternative text from the playbook
   - "comment_only": The playbook doesn't specify exact replacement text, so only add a comment
5. If action is "reject_restore" or "reject_replace", provide the EXACT text to insert as "auto_redline_text"
6. Determine the risk level (Low/Medium/High)
7. Create a clear, actionable comment explaining the reasoning

CRITICAL - PLAYBOOK REFERENCE REQUIRED:
- You MUST cite the specific playbook clause FIRST before making any recommendation
- The "playbook_principle" field MUST contain the EXACT quoted text from the playbook above
- Include the clause number/name (e.g., "CLAUSE 6: TERM AND TERMINATION - Standard Position: 3-year term...")
- If no specific playbook clause applies, state "No specific playbook guidance found for this change"
- All assessments and actions MUST be justified by the cited playbook principle

CRITICAL FOR AUTO-REDLINING:
- If the playbook specifies exact language (e.g., "3-year term"), use that exact language in auto_redline_text
- For DELETIONS: If rejecting, set action to "reject_restore" and auto_redline_text to the deleted text that should be restored
- For INSERTIONS: If rejecting, set action to "reject_replace" with empty auto_redline_text (the insertion will be removed)
- For REPLACEMENTS: If rejecting, set action to "reject_restore" to restore old text, or "reject_replace" with playbook-specified alternative
- If the playbook has fallback ranges (e.g., "1 to 5 years acceptable"), and the change falls within that range, use "accept"
- Only use "comment_only" when the playbook doesn't provide specific language and you cannot determine exact replacement text

Format your response as JSON with this structure:
{{
  "analyses": [
    {{
      "redline_number": 1,
      "playbook_principle": "Exact text of the playbook principle that applies to this redline",
      "assessment": "Detailed assessment of how the change aligns with the playbook",
      "auto_redline_action": "accept|reject_restore|reject_replace|comment_only",
      "auto_redline_text": "Exact text to insert as counter-redline (empty string if accept or comment_only, or if rejecting an insertion)",
      "response": "Recommended actions based on the playbook principle - be specific about what should be done",
      "fallbacks": "Recommended fallback positions or alternative approaches if the primary recommendation cannot be accepted.",
      "risk_level": "Low|Medium|High",
      "comment_text": "Inline comment explaining the auto-redline action and reasoning (clear, actionable guidance for the reviewer)"
    }}
  ]
}}

EXAMPLES:
- Counterparty changes "three (3) years" to "5 years": If playbook says 3-year term is standard but 1-5 years is acceptable fallback, use "accept" (within fallback range)
- Counterparty deletes Clause 4(d) about AI training: If playbook says this is critical and cannot be deleted, use "reject_restore" with the deleted clause text
- Counterparty changes governing law from California to Michigan: If playbook says only CA, DE, or NY acceptable, use "reject_replace" with "California" as auto_redline_text
- Counterparty adds new indemnification language: If playbook says no indemnification allowed but doesn't specify replacement, use "reject_replace" with empty auto_redline_text to remove it

Additional context: {context or 'None provided'}
"""
        return prompt
    
    def _call_ai(self, prompt: str) -> str:
        """Call AI API and return response."""
        if self.provider == "openai":
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a legal technology assistant specializing in contract review and redline analysis."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                response_format={"type": "json_object"} if "gpt-4" in self.model.lower() else None
            )
            return response.choices[0].message.content
        elif self.provider == "anthropic":
            # Anthropic models support JSON mode in newer versions
            response = self.client.messages.create(
                model=self.model,
                max_tokens=4000,
                system="You are a legal technology assistant specializing in contract review and redline analysis. Always respond with valid JSON.",
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            return response.content[0].text
    
    def _parse_ai_response(self, ai_response: str, redlines: List[Dict], redline_number: int = 1) -> List[Dict]:
        """Parse AI response and match to redlines.
        
        Args:
            ai_response: The raw AI response text
            redlines: List of redlines being analyzed
            redline_number: The number of the redline being analyzed (for single redline analysis)
        """
        import json
        import re
        
        try:
            # Try to extract JSON from response
            # Handle cases where AI wraps JSON in markdown code blocks
            response_text = ai_response.strip()
            
            # Remove markdown code blocks if present
            if response_text.startswith('```'):
                # Extract JSON from code block
                lines = response_text.split('\n')
                json_start = None
                json_end = None
                for i, line in enumerate(lines):
                    if line.strip().startswith('```'):
                        if json_start is None:
                            json_start = i + 1
                        else:
                            json_end = i
                            break
                if json_start and json_end:
                    response_text = '\n'.join(lines[json_start:json_end])
            
            # Try to find JSON object in the response
            # Look for { ... } pattern
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            if json_match:
                response_text = json_match.group(0)
            
            data = json.loads(response_text)
            
            # Handle both array format and single object format
            analyses = []
            if 'analyses' in data:
                analyses = data.get('analyses', [])
            elif isinstance(data, list):
                analyses = data
            elif isinstance(data, dict) and 'redline_number' in data:
                # Single analysis object
                analyses = [data]
            elif isinstance(data, dict):
                # Try to treat the whole object as a single analysis
                analyses = [data]
            
            # Match analyses to redlines
            results = []
            for analysis in analyses:
                # For single redline analysis, use the provided redline_number
                # Otherwise, use the redline_number from the analysis
                if len(redlines) == 1:
                    redline_idx = 0
                else:
                    redline_num = analysis.get('redline_number', redline_number) - 1
                    redline_idx = redline_num if 0 <= redline_num < len(redlines) else 0
                
                if 0 <= redline_idx < len(redlines):
                    results.append({
                        'redline': redlines[redline_idx],
                        'playbook_principle': analysis.get('playbook_principle', ''),
                        'assessment': analysis.get('assessment', ''),
                        'response': analysis.get('response', ''),
                        'fallbacks': analysis.get('fallbacks', ''),
                        'risk_level': analysis.get('risk_level', 'Medium'),
                        'comment_text': analysis.get('comment_text', ''),
                        'auto_redline_action': analysis.get('auto_redline_action', 'comment_only'),
                        'auto_redline_text': analysis.get('auto_redline_text', '')
                    })
            
            return results if results else None
        
        except json.JSONDecodeError as e:
            print(f"    JSON decode error: {e}")
            print(f"    Response text (first 500 chars): {ai_response[:500]}")
            # Fallback: create simple responses
            results = []
            for redline in redlines:
                results.append({
                    'redline': redline,
                    'playbook_principle': 'General Legal Guidelines',
                    'assessment': 'Requires review - JSON parsing failed',
                    'response': ai_response[:200] if ai_response else 'Please review manually',
                    'fallbacks': '',
                    'risk_level': 'Medium',
                    'comment_text': 'Please review this change against legal playbook.',
                    'auto_redline_action': 'comment_only',
                    'auto_redline_text': ''
                })
            return results
        except Exception as e:
            print(f"    Unexpected error parsing response: {type(e).__name__}: {e}")
            import traceback
            print(f"    Traceback: {traceback.format_exc()}")
            return None

