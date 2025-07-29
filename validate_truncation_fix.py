#!/usr/bin/env python3
"""
Quick validation script to verify the truncation fix is working.
This can be run in production environments to confirm the fix.
"""

def validate_fix():
    """Quick validation that the truncation fix is properly implemented."""
    try:
        from utils import get_max_tokens_for_step
        
        # Test the critical steps that were causing truncation
        concepts_tokens = get_max_tokens_for_step('step5')
        model_tokens = get_max_tokens_for_step('step6')
        
        print("🔍 Validating Truncation Fix...")
        print(f"Concepts (step5): {concepts_tokens} tokens")
        print(f"Conceptual Model (step6): {model_tokens} tokens")
        
        if concepts_tokens >= 3000 and model_tokens >= 4000:
            print("✅ PASS: Token limits are sufficient to prevent truncation")
            print("✅ Concepts and conceptual model outputs should now be complete")
            return True
        else:
            print("❌ FAIL: Token limits are still too low")
            return False
            
    except ImportError:
        print("❌ FAIL: Cannot import get_max_tokens_for_step function")
        return False
    except Exception as e:
        print(f"❌ FAIL: Unexpected error: {e}")
        return False

if __name__ == "__main__":
    success = validate_fix()
    if success:
        print("\n🎉 Truncation fix validation PASSED!")
    else:
        print("\n❌ Truncation fix validation FAILED!")
    exit(0 if success else 1)