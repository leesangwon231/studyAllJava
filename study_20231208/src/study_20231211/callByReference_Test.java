package study_20231211;

public class callByReference_Test {

	/**
	 * 2023-12-11
	 * java에는 포인터가 없다
	 * +) 그럴 수 밖에 없는 이유는 자바의 특징인 garbege collector 때문이라는데 자세한건 더 공부하는 걸로 하고
	 * +) 포인터(Pointer), 참조(Reference) 차이 확인해보기 -> [간략] 포인터는 메모리를 직접 핸들링할 수 있고 참조는 불가능
	 * 
	 * 해당 특징을 공부하기 위해 2가지로 확인
	 * 1. 함수에 변수를 전달하고 값 변경 후 리턴없이 확인 => 당연히 변경되지 않음
	 * 2. class 객체를 생성한 후 객체를 함수에 전달해서 값 변경 후 리턴없이 확인 => 함수에서 변경한 값으로 변경됨.
	 * 
	 * 
	 * 
	 * 
	 */
	
	
	public static void main(String[] args) {
		
		String comeBack = "hello!";
		
		System.out.println("before : " + comeBack);
		
		testFunction(comeBack);
		
		System.out.println("after : " + comeBack);
		
		ReferenceTest test01 = new ReferenceTest(0, "Hi!");
		
		System.out.println("before ReferenceTest :" + test01.getNumber() + " , " + test01.getText());
		
		changeValue(test01, test01.getNumber(), test01.getText());
		
		System.out.println("after ReferenceTest :" + test01.getNumber() + " , " + test01.getText());
		

	}

	
	public static void testFunction(String test) {
		test = "call by reference";
	}
	
	public static void changeValue(ReferenceTest sample, int n, String t) {
		sample.setNumber(100);
		sample.setText("Hello!");
		
		n = 999;
		t = "Error";
	}
	
	public static class ReferenceTest {
		
		private int number;
		private String text;
		
		public ReferenceTest(int n, String t){
			this.number = n;
			this.text = t;
		}
		
		public String getText() {
			return text;
		}

		public void setText(String text) {
			this.text = text;
		}


		public int getNumber() {
			return number;
		}

		public void setNumber(int number) {
			this.number = number;
		}

	}
	
}


