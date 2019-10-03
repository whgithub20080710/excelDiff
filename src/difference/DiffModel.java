package difference;

public class DiffModel {

	private String self;
	
	private String other;
	
	private String selfAmount;
	
	private String otherAmount;

	public String getSelf() {
		return self;
	}

	public void setSelf(String self) {
		this.self = self;
	}

	public String getOther() {
		return other;
	}

	public void setOther(String other) {
		this.other = other;
	}

	public String getSelfAmount() {
		return selfAmount;
	}

	public void setSelfAmount(String selfAmount) {
		this.selfAmount = selfAmount;
	}

	public String getOtherAmount() {
		return otherAmount;
	}

	public void setOtherAmount(String otherAmount) {
		this.otherAmount = otherAmount;
	}

	public DiffModel(String self, String selfAmount, String other, String otherAmount) {
		super();
		this.self = self;
		this.other = other;
		this.selfAmount = selfAmount;
		this.otherAmount = otherAmount;
	}

	public DiffModel() {
		super();
		// TODO Auto-generated constructor stub
	}

	@Override
	public String toString() {
		return ("".equals(self)?"null":self)+","+("".equals(selfAmount)?"null":selfAmount)+","+("".equals(other)?"null":other)+","+("".equals(otherAmount)?"null":otherAmount);
	}
	
	
	
}
