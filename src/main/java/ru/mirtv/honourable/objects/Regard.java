/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.mirtv.honourable.objects;

/**
 *
 * @author Babikov_PV
 */
public class Regard {

  private String name;
  private String post;
  private String reason;

  public Regard(String name, String post, String reason) {
    if (name != "" && post != "" && reason != "") {
      this.name = name;
      this.post = post;
      this.reason = reason;
    }
  }

  /**
   * @return the name
   */
  public String getName() {
    return name;
  }

  /**
   * @param name the name to set
   */
  public void setName(String name) {
    this.name = name;
  }

  /**
   * @return the post
   */
  public String getPost() {
    return post;
  }

  /**
   * @param post the post to set
   */
  public void setPost(String post) {
    this.post = post;
  }

  /**
   * @return the reason
   */
  public String getReason() {
    return reason;
  }

  /**
   * @param reason the reason to set
   */
  public void setReason(String reason) {
    this.reason = reason;
  }
}
